import type { SPFI } from '@pnp/sp';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../../pnpjsConfig';
import { getLatestItemByCanId } from './sharePointListUtils';

export const DEFAULT_HR_DOCUMENT_LIBRARY = 'Link Document Library';

type ListArchiveConfig = {
  listTitle: string;
  /** Prefix for uploaded file names, e.g. JoiningFormalities */
  docTypePrefix: string;
};

const LISTS_WITH_ATTACHMENTS: ListArchiveConfig[] = [
  { listTitle: 'JoiningFormalities', docTypePrefix: 'JoiningFormalities' },
  { listTitle: 'InsuranceNomination', docTypePrefix: 'TeamLifeInsurance' },
  { listTitle: 'PersonalAccidentalScheme', docTypePrefix: 'PersonalAccidentInsurance' },
  { listTitle: 'GratuityNomination', docTypePrefix: 'GratuityNomination' }
];

function sanitizeFolderName(name: string): string {
  return name.replace(/[~"#%&*:<>?/\\{|}]/g, '_').trim();
}

function sanitizeFileName(name: string): string {
  return name.replace(/[~"#%&*:<>?/\\{|}]/g, '_');
}

async function ensureEmployeeFolder(
  sp: SPFI,
  libraryTitle: string,
  folderName: string
): Promise<string> {
  const lib = sp.web.lists.getByTitle(libraryTitle);
  const libInfo = await lib.select('RootFolder/ServerRelativeUrl')();
  const rootUrl = libInfo.RootFolder?.ServerRelativeUrl;

  if (!rootUrl) {
    throw new Error(`Could not resolve root folder for library "${libraryTitle}".`);
  }

  const folderUrl = `${rootUrl}/${folderName}`;

  try {
    await sp.web.folders.addUsingPath(folderUrl, true);
  } catch {
    // Folder may already exist.
  }

  return folderUrl;
}

async function copyListAttachments(
  sp: SPFI,
  listTitle: string,
  itemId: number,
  docTypePrefix: string,
  folderUrl: string
): Promise<number> {
  let copied = 0;

  const item = sp.web.lists.getByTitle(listTitle).items.getById(itemId);
  const attachments = await item.attachmentFiles();

  for (const attachment of attachments) {
    const blob = await sp.web
      .getFileByServerRelativePath(attachment.ServerRelativeUrl)
      .getBlob();

    const targetName = sanitizeFileName(`${docTypePrefix}_${attachment.FileName}`);

    await sp.web
      .getFolderByServerRelativePath(folderUrl)
      .files.addUsingPath(targetName, blob, { Overwrite: true });

    copied += 1;
  }

  return copied;
}

/**
 * After HR verification, copy list attachments into
 * {Document Library}/{EmployeeID}/{DocType}_{FileName}.
 */
export async function archiveEmployeeDocumentsToLibrary(
  context: WebPartContext,
  canId: string | number,
  documentLibraryTitle: string = DEFAULT_HR_DOCUMENT_LIBRARY
): Promise<{ folderName: string; filesCopied: number }> {
  const sp = getSP(context);
  const joining = await getLatestItemByCanId(sp, 'JoiningFormalities', canId);
  const employeeId = String(
    joining?.hr_employee_ID || joining?.employeeId || canId
  ).trim();
  const folderName = sanitizeFolderName(employeeId || String(canId));

  const folderUrl = await ensureEmployeeFolder(sp, documentLibraryTitle, folderName);

  let filesCopied = 0;

  for (const config of LISTS_WITH_ATTACHMENTS) {
    const row = await getLatestItemByCanId(sp, config.listTitle, canId);
    const itemId = Number(row?.Id);

    if (!itemId) {
      continue;
    }

    filesCopied += await copyListAttachments(
      sp,
      config.listTitle,
      itemId,
      config.docTypePrefix,
      folderUrl
    );
  }

  return { folderName, filesCopied };
}
