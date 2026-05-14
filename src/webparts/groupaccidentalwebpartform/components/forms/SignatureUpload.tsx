import * as React from 'react';

type SignatureUploadProps = {
  name: string;
  value: string;
  onChange: (value: string) => void;
  onBlur?: React.FocusEventHandler<HTMLInputElement>;
};

const SignatureUpload = ({
  name,
  value,
  onChange,
  onBlur
}: SignatureUploadProps): JSX.Element => {
  const inputRef = React.useRef<HTMLInputElement>(null);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.currentTarget.files?.[0];

    if (!file) {
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      onChange(typeof reader.result === 'string' ? reader.result : '');
    };
    reader.readAsDataURL(file);
  };

  return (
    <div>
      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        name={name}
        onChange={handleFileChange}
        onBlur={onBlur}
        style={{ display: 'none' }}
        data-pdf-hide="true"
      />

      <div
        role="button"
        tabIndex={0}
        onClick={() => inputRef.current?.click()}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            inputRef.current?.click();
          }
        }}
        style={{
          minHeight: '120px',
          height: '120px',
          border: '1px solid #ced4da',
          borderRadius: '0.375rem',
          background: '#fff',
          display: 'flex',
          alignItems: 'center',
          justifyContent: value ? 'center' : 'flex-start',
          overflow: 'hidden',
          padding: '0.75rem',
          cursor: 'pointer'
        }}
      >
        {value ? (
          <img
            src={value}
            alt="Signature preview"
            style={{
              maxWidth: '100%',
              maxHeight: '110px',
              objectFit: 'contain'
            }}
          />
        ) : (
          <span className="text-muted" style={{ fontSize: '14px' }}>Upload signature</span>
        )}
      </div>
    </div>
  );
};

export default SignatureUpload;
