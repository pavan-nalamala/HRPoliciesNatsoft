import * as React from 'react';
import * as Yup from 'yup';
import moment from 'moment';

export const DATE_INPUT_FORMAT = 'DD/MM/YYYY';
export const DATE_INPUT_PLACEHOLDER = 'DD/MM/YYYY';
const NATIVE_DATE_FORMAT = 'YYYY-MM-DD';

export const formatDateInput = (value: string): string => {
  const digitsOnly = value.replace(/\D/g, '').slice(0, 8);

  if (digitsOnly.length <= 2) {
    return digitsOnly;
  }

  if (digitsOnly.length <= 4) {
    return `${digitsOnly.slice(0, 2)}/${digitsOnly.slice(2)}`;
  }

  return `${digitsOnly.slice(0, 2)}/${digitsOnly.slice(2, 4)}/${digitsOnly.slice(4)}`;
};

export const createDateValueChangeHandler = (
  setFieldValue: (field: string, value: string) => Promise<unknown>,
  fieldName: string
) => (value: string): void => {
  setFieldValue(fieldName, formatDateInput(value)).catch(() => undefined);
};

/** Employee DOB field — keeps DD/MM/YYYY and syncs shared state across tabs. */
export const createSharedDobChangeHandler = (
  setFieldValue: (field: string, value: string) => Promise<unknown>,
  fieldName: string,
  onSharedDateOfBirthChange?: (value: string) => void
) => (value: string): void => {
  const formatted = formatDateInput(value);
  setFieldValue(fieldName, formatted).catch(() => undefined);
  onSharedDateOfBirthChange?.(formatted);
};

export const formatDateForDisplay = (date: Date): string => {
  return moment(date).format(DATE_INPUT_FORMAT);
};

export type SharePointDateFieldKind = 'date' | 'datetime';

function parseFormDateValue(
  value: string | undefined | null
): moment.Moment | null {
  if (!value?.trim()) {
    return null;
  }

  const trimmed = value.trim();
  const strictUi = moment(trimmed, DATE_INPUT_FORMAT, true);
  if (strictUi.isValid()) {
    return strictUi;
  }

  const isoDate = moment(trimmed, 'YYYY-MM-DD', true);
  if (isoDate.isValid()) {
    return isoDate;
  }

  const fromStorage = moment(trimmed);
  return fromStorage.isValid() ? fromStorage : null;
}

/** Parse form date for SharePoint (date-only or DateTime column). */
export function parseFormDateForSharePoint(
  value: string | undefined | null,
  kind: SharePointDateFieldKind = 'datetime'
): string | null {
  const parsed = parseFormDateValue(value);
  if (!parsed) {
    return null;
  }

  if (kind === 'date') {
    return parsed.format('YYYY-MM-DD');
  }

  return toSharePointDateTimeIso(parsed);
}

/** DOB / birth-date columns — SharePoint DateTime fields (date-only display). */
export function parseDobForSharePoint(
  value: string | undefined | null
): string | null {
  return parseFormDateForSharePoint(value, 'datetime');
}

function toSharePointDateTimeIso(parsed: moment.Moment): string {
  return moment
    .utc({
      year: parsed.year(),
      month: parsed.month(),
      date: parsed.date()
    })
    .toISOString();
}

/** @deprecated Use parseFormDateForSharePoint(value, 'datetime') */
export function toSharePointDateTime(
  value: string | undefined | null,
  inputFormat: string = DATE_INPUT_FORMAT
): string | null {
  void inputFormat;
  return parseFormDateForSharePoint(value, 'datetime');
}

/** Format stored SharePoint date values for display in form fields (DD/MM/YYYY). */
export function formatSharePointDateForInput(
  value: string | Date | undefined | null
): string {
  if (value === undefined || value === null || value === '') {
    return '';
  }

  const raw = typeof value === 'string' ? value.trim() : value;
  const utcParsed = moment.utc(raw);
  if (utcParsed.isValid()) {
    return utcParsed.format(DATE_INPUT_FORMAT);
  }

  const localParsed = moment(raw);
  return localParsed.isValid() ? localParsed.format(DATE_INPUT_FORMAT) : '';
}

const formatDateForNativePicker = (value: string): string => {
  if (!value) {
    return '';
  }

  const parsedValue = moment(value, DATE_INPUT_FORMAT, true);
  return parsedValue.isValid() ? parsedValue.format(NATIVE_DATE_FORMAT) : '';
};

const formatNativePickerValue = (value: string): string => {
  if (!value) {
    return '';
  }

  return moment(value, NATIVE_DATE_FORMAT, true).format(DATE_INPUT_FORMAT);
};

type DatePickerInputProps = {
  name: string;
  value: string;
  onValueChange: (value: string) => void;
  onBlur?: React.FocusEventHandler<HTMLInputElement>;
  className?: string;
  style?: React.CSSProperties;
  disabled?: boolean;
};

export const DatePickerInput = ({
  name,
  value,
  onValueChange,
  onBlur,
  className,
  style,
  disabled
}: DatePickerInputProps): JSX.Element => {
  const nativeInputRef = React.useRef<HTMLInputElement>(null);

  const handleOpenPicker = (): void => {
    const dateInput = nativeInputRef.current as HTMLInputElement & {
      showPicker?: () => void;
    };

    if (dateInput?.showPicker) {
      dateInput.showPicker();
      return;
    }

    dateInput?.click();
  };

  return (
    <div className="input-group">
      <input
        type="text"
        name={name}
        className={className ?? 'form-control'}
        style={style}
        placeholder={DATE_INPUT_PLACEHOLDER}
        value={value}
        onChange={(event) => onValueChange(event.target.value)}
        onBlur={onBlur}
        disabled={disabled}
      />
      <button
        type="button"
        className="btn btn-outline-secondary"
        onClick={handleOpenPicker}
        disabled={disabled}
        aria-label={`Open calendar for ${name}`}
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="16"
          height="16"
          fill="currentColor"
          viewBox="0 0 16 16"
          aria-hidden="true"
        >
          <path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5ZM1 5v8a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V5H1Zm1-3a1 1 0 0 0-1 1v1h14V3a1 1 0 0 0-1-1H2Z" />
        </svg>
      </button>
      <input
        ref={nativeInputRef}
        type="date"
        tabIndex={-1}
        aria-hidden="true"
        value={formatDateForNativePicker(value)}
        onChange={(event) => onValueChange(formatNativePickerValue(event.target.value))}
        style={{
          position: 'absolute',
          width: 0,
          height: 0,
          opacity: 0,
          pointerEvents: 'none'
        }}
      />
    </div>
  );
};

export const createDateValidation = (requiredMessage: string, invalidMessage: string): Yup.StringSchema<string> => {
  return Yup.string()
    .required(requiredMessage)
    .test('valid-date', invalidMessage, (value) => {
      if (!value) {
        return false;
      }

      return moment(value, DATE_INPUT_FORMAT, true).isValid();
    });
};

export const createMatchingDateValidation = (
  schema: Yup.StringSchema<string>,
  expectedDate: string,
  message: string
): Yup.StringSchema<string> => {
  return schema.test('matching-date', message, (value) => {
    if (!value || !expectedDate) {
      return true;
    }

    const parsedValue = moment(value, DATE_INPUT_FORMAT, true);
    const normalizedExpected =
      formatSharePointDateForInput(expectedDate) || expectedDate;
    const parsedExpectedDate = moment(normalizedExpected, DATE_INPUT_FORMAT, true);

    if (!parsedValue.isValid() || !parsedExpectedDate.isValid()) {
      return true;
    }

    return parsedValue.isSame(parsedExpectedDate, 'day');
  });
};
