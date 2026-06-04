/** Values stored in EPFDeclarationForm11.HRStatus (and shown in the UI). */
export const HR_STATUS = {
  PENDING_FROM_HR: 'Pending from HR',
  VERIFIED_COMPLETED: 'Verified and Completed'
} as const;

export type HrStatusValue =
  | typeof HR_STATUS.PENDING_FROM_HR
  | typeof HR_STATUS.VERIFIED_COMPLETED
  | '';

/** Map legacy list values to current status labels. */
export function normalizeHrStatus(status?: string | null): string {
  if (!status) {
    return '';
  }

  const trimmed = status.trim();

  if (trimmed === 'Pending' || trimmed === HR_STATUS.PENDING_FROM_HR) {
    return HR_STATUS.PENDING_FROM_HR;
  }

  if (trimmed === 'Completed' || trimmed === HR_STATUS.VERIFIED_COMPLETED) {
    return HR_STATUS.VERIFIED_COMPLETED;
  }

  return trimmed;
}

/** Employee final submit done — forms are view-only for the employee. */
export function isEmployeeSubmissionLocked(hrStatus?: string | null): boolean {
  const normalized = normalizeHrStatus(hrStatus);
  return (
    normalized === HR_STATUS.PENDING_FROM_HR ||
    normalized === HR_STATUS.VERIFIED_COMPLETED
  );
}

export function isHrVerificationComplete(hrStatus?: string | null): boolean {
  return normalizeHrStatus(hrStatus) === HR_STATUS.VERIFIED_COMPLETED;
}

export function getHrStatusDisplayLabel(hrStatus?: string | null): string {
  const normalized = normalizeHrStatus(hrStatus);

  if (!normalized) {
    return 'In Progress';
  }

  return normalized;
}

export type WorkflowBadgeKind = 'inProgress' | 'pendingHr' | 'verified';

export function getWorkflowBadgeKind(
  hrStatus?: string | null,
  isFinallySubmitted?: boolean
): WorkflowBadgeKind {
  const normalized = normalizeHrStatus(hrStatus);

  if (normalized === HR_STATUS.VERIFIED_COMPLETED) {
    return 'verified';
  }

  if (isFinallySubmitted || normalized === HR_STATUS.PENDING_FROM_HR) {
    return 'pendingHr';
  }

  return 'inProgress';
}
