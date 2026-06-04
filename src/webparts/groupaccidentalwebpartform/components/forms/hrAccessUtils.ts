/**
 * HR admin accounts — full edit access, HR tabs, deep links, employer PF section.
 * Add new HR emails here (lowercase).
 */
export const HR_ADMIN_EMAILS: readonly string[] = [
  'hr@nat.in',
  'hr@natit.in',
  'testgauge@natit.in'
];

/** Normalize email / UPN from PnP profile or SharePoint claims. */
export function normalizeUserEmail(email?: string | null): string {
  if (!email) {
    return '';
  }

  let value = email.trim().toLowerCase();
  const membershipPrefix = 'i:0#.f|membership|';

  if (value.startsWith(membershipPrefix)) {
    value = value.substring(membershipPrefix.length);
  }

  return value.trim();
}

/** Production form page (NatIt HR Recruitment site). */
export const DEFAULT_HR_FORM_PAGE_PATH =
  '/sites/NatIt_HRRecruitment/SitePages/GroupPersonalAccidentalNomination.aspx';

export const DEFAULT_SHAREPOINT_ORIGIN = 'https://natitin.sharepoint.com';

export const HR_TRACKING_QUERY_PARAM = 'trackingId';

export function isHrUserEmail(email?: string | null): boolean {
  const normalized = normalizeUserEmail(email);
  if (!normalized) {
    return false;
  }

  return HR_ADMIN_EMAILS.indexOf(normalized) !== -1;
}

/** Token from `#/admin-edit/{token}` in the page URL hash. */
export function getAdminDeepLinkToken(): string | null {
  const hash = window.location.hash || '';
  const match = hash.match(/#\/admin-edit\/([^/?]+)/i);
  return match?.[1] ?? null;
}

/**
 * Candidate / employee form ID from the URL (maps to Candidate Information list ID / can_id).
 * Prefer trackingId to match production links.
 */
export function getHrReviewCandidateId(): string | null {
  const params = new URLSearchParams(window.location.search);
  return (
    params.get(HR_TRACKING_QUERY_PARAM) ||
    params.get('canId') ||
    params.get('candidateId') ||
    params.get('employeeFormId') ||
    params.get('employeeId')
  );
}

export function isHrDeepLinkMode(): boolean {
  const params = new URLSearchParams(window.location.search);
  return params.get('mode') === 'hr';
}

export type HrReviewLinkOptions = {
  /** e.g. https://natitin.sharepoint.com */
  origin?: string;
  /** e.g. /sites/NatIt_HRRecruitment/SitePages/GroupPersonalAccidentalNomination.aspx */
  formPagePath?: string;
  /** Include mode=hr (required for HR review UI). Default true. */
  includeHrMode?: boolean;
};

/**
 * HR review URL — production shape:
 * https://natitin.sharepoint.com/sites/NatIt_HRRecruitment/SitePages/GroupPersonalAccidentalNomination.aspx?mode=hr&trackingId=65
 */
export function buildHrReviewLinkByCandidateId(
  candidateId: string | number,
  options: HrReviewLinkOptions = {}
): string {
  const origin = options.origin || DEFAULT_SHAREPOINT_ORIGIN;
  const formPagePath = options.formPagePath || DEFAULT_HR_FORM_PAGE_PATH;
  const includeHrMode = options.includeHrMode !== false;
  const id = encodeURIComponent(String(candidateId));

  const params = new URLSearchParams();
  if (includeHrMode) {
    params.set('mode', 'hr');
  }
  params.set(HR_TRACKING_QUERY_PARAM, id);

  return `${origin}${formPagePath}?${params.toString()}`;
}

/** Server-relative path of the page hosting the web part. */
export function getFormPageServerRelativePath(): string {
  return `${window.location.pathname}${window.location.search}`;
}

/** Build HR deep link using the current SharePoint page (legacy token links). */
export function buildAdminDeepLinkFromCurrentPage(token: string): string {
  return `${window.location.origin}${getFormPageServerRelativePath()}#/admin-edit/${token}?mode=hr`;
}

export function buildAdminDeepLink(
  origin: string,
  serverRelativeWebUrl: string,
  pagePath: string,
  token: string
): string {
  const base = `${origin}${serverRelativeWebUrl}${pagePath}`;
  return `${base}#/admin-edit/${token}?mode=hr`;
}

/**
 * Fix legacy AdminDeepLink values that pointed at a non-existent page.
 */
export function repairAdminDeepLinkIfNeeded(
  storedLink: string | undefined,
  token: string,
  formPagePath?: string
): string {
  if (!storedLink) {
    return buildHrReviewLinkByCandidateId(token, { formPagePath });
  }

  if (
    storedLink.includes('/SitePages/PFForms.aspx') ||
    storedLink.includes('canId=')
  ) {
    const canIdMatch = storedLink.match(/[?&]canId=([^&]+)/i);
    const id = canIdMatch?.[1] || token;
    return buildHrReviewLinkByCandidateId(id, { formPagePath });
  }

  return storedLink;
}
