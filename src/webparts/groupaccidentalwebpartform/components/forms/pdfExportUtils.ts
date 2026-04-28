declare global {
  interface Window {
    html2canvas?: (
      element: HTMLElement,
      options?: Record<string, unknown>
    ) => Promise<HTMLCanvasElement>;
    jspdf?: {
      jsPDF: new (options?: Record<string, unknown>) => {
        internal: {
          pageSize: {
            getWidth: () => number;
            getHeight: () => number;
          };
        };
        addImage: (
          imageData: string,
          format: string,
          x: number,
          y: number,
          width: number,
          height: number
        ) => void;
        addPage: () => void;
        save: (fileName: string) => void;
      };
    };
  }
}

const HTML2CANVAS_CDN =
  'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
const JSPDF_CDN =
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';

const loadScript = (src: string): Promise<void> =>
  new Promise((resolve, reject) => {
    const existingScript = document.querySelector<HTMLScriptElement>(
      `script[src="${src}"]`
    );

    if (existingScript) {
      if (existingScript.dataset.loaded === 'true') {
        resolve();
        return;
      }

      existingScript.addEventListener('load', () => resolve(), { once: true });
      existingScript.addEventListener(
        'error',
        () => reject(new Error(`Failed to load script: ${src}`)),
        { once: true }
      );
      return;
    }

    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.addEventListener(
      'load',
      () => {
        script.dataset.loaded = 'true';
        resolve();
      },
      { once: true }
    );
    script.addEventListener(
      'error',
      () => reject(new Error(`Failed to load script: ${src}`)),
      { once: true }
    );
    document.head.appendChild(script);
  });

const ensurePdfLibraries = async (): Promise<void> => {
  if (!window.html2canvas) {
    await loadScript(HTML2CANVAS_CDN);
  }

  if (!window.jspdf?.jsPDF) {
    await loadScript(JSPDF_CDN);
  }
};

const createTextValueNode = (
  ownerDocument: Document,
  value: string
): HTMLDivElement => {
  const node = ownerDocument.createElement('div');
  node.textContent = value || ' ';
  node.style.minHeight = '38px';
  node.style.padding = '0.375rem 0.75rem';
  node.style.whiteSpace = 'pre-wrap';
  node.style.wordBreak = 'break-word';
  node.style.border = 'none';
  node.style.background = 'transparent';
  return node;
};

const prepareCloneForPdf = (sourceElement: HTMLElement): HTMLElement => {
  const clonedElement = sourceElement.cloneNode(true) as HTMLElement;

  clonedElement
    .querySelectorAll<HTMLElement>('button, [data-pdf-hide="true"]')
    .forEach((element) => element.remove());

  clonedElement.querySelectorAll<HTMLElement>('input, textarea, select').forEach((element) => {
    if (element instanceof HTMLInputElement) {
      if (element.type === 'hidden') {
        element.remove();
        return;
      }

      if (element.type === 'checkbox' || element.type === 'radio') {
        element.disabled = true;
        return;
      }

      const textNode = createTextValueNode(element.ownerDocument, element.value);
      element.replaceWith(textNode);
      return;
    }

    if (element instanceof HTMLTextAreaElement) {
      const textNode = createTextValueNode(element.ownerDocument, element.value);
      textNode.style.minHeight = '72px';
      element.replaceWith(textNode);
      return;
    }

    if (element instanceof HTMLSelectElement) {
      const selectedText = element.options[element.selectedIndex]?.text ?? '';
      const textNode = createTextValueNode(element.ownerDocument, selectedText);
      element.replaceWith(textNode);
    }
  });

  const styleTag = clonedElement.ownerDocument.createElement('style');
  styleTag.textContent = `
    * {
      box-shadow: none !important;
    }

    iframe {
      display: none !important;
    }

    .card,
    .table,
    .table td,
    .table th,
    .border,
    .rounded,
    .rounded-4 {
      border-color: #d7d7d7 !important;
    }
  `;
  clonedElement.prepend(styleTag);

  return clonedElement;
};

const openPrintFallback = (
  sourceElement: HTMLElement,
  fileName: string
): void => {
  const printWindow = window.open('', '_blank', 'noopener,noreferrer,width=1024,height=768');

  if (!printWindow) {
    throw new Error('Unable to open print window for PDF fallback.');
  }

  const printableClone = prepareCloneForPdf(sourceElement);

  printWindow.document.open();
  printWindow.document.write(`
    <!DOCTYPE html>
    <html>
      <head>
        <title>${fileName}</title>
        <meta charset="utf-8" />
        <style>
          body {
            margin: 0;
            padding: 24px;
            font-family: Arial, sans-serif;
            background: #fff;
          }

          img {
            max-width: 100%;
          }
        </style>
      </head>
      <body></body>
    </html>
  `);
  printWindow.document.body.appendChild(printableClone);
  printWindow.document.close();
  printWindow.focus();

  window.setTimeout(() => {
    printWindow.print();
  }, 400);
};

export const downloadElementAsPdf = async (
  sourceElement: HTMLElement,
  fileName: string
): Promise<void> => {
  let pdfContainer: HTMLDivElement | null = null;

  try {
    await ensurePdfLibraries();

    if (!window.html2canvas || !window.jspdf?.jsPDF) {
      throw new Error('PDF libraries are not available.');
    }

    pdfContainer = document.createElement('div');
    pdfContainer.style.position = 'fixed';
    pdfContainer.style.left = '-20000px';
    pdfContainer.style.top = '0';
    pdfContainer.style.width = `${Math.max(sourceElement.scrollWidth, 900)}px`;
    pdfContainer.style.padding = '24px';
    pdfContainer.style.background = '#ffffff';
    pdfContainer.style.zIndex = '-1';

    const printableClone = prepareCloneForPdf(sourceElement);
    pdfContainer.appendChild(printableClone);
    document.body.appendChild(pdfContainer);

    const canvas = await window.html2canvas(pdfContainer, {
      backgroundColor: '#ffffff',
      scale: 2,
      useCORS: true,
      logging: false
    });

    const imageData = canvas.toDataURL('image/png');
    const pdf = new window.jspdf.jsPDF({
      orientation: 'p',
      unit: 'mm',
      format: 'a4'
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const imageWidth = pageWidth - 20;
    const imageHeight = (canvas.height * imageWidth) / canvas.width;

    let renderedHeight = imageHeight;
    let yPosition = 10;

    pdf.addImage(imageData, 'PNG', 10, yPosition, imageWidth, imageHeight);
    renderedHeight -= pageHeight - 20;

    while (renderedHeight > 0) {
      yPosition = renderedHeight - imageHeight + 10;
      pdf.addPage();
      pdf.addImage(imageData, 'PNG', 10, yPosition, imageWidth, imageHeight);
      renderedHeight -= pageHeight - 20;
    }

    pdf.save(fileName);
  } catch (error) {
    console.error('PDF export failed, opening print fallback instead.', error);
    openPrintFallback(sourceElement, fileName);
  } finally {
    if (pdfContainer?.parentNode) {
      pdfContainer.parentNode.removeChild(pdfContainer);
    }
  }
};

export const downloadFileFromUrl = (fileUrl: string, fileName: string): void => {
  const link = document.createElement('a');
  link.href = fileUrl;
  link.download = fileName;
  link.target = '_blank';
  link.rel = 'noopener noreferrer';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};
