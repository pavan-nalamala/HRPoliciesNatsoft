/** HR admin accounts allowed to use admin deep links and HR-only form sections. */
export const HR_ADMIN_EMAILS: readonly string[] = [
  'hr@nat.in',
  'hr@natit.in',
  'testgauge@natit.in'
];

export function isHrUserEmail(email?: string | null): boolean {
  if (!email) {
    return false;
  }

  const normalized = email.trim().toLowerCase();
  return HR_ADMIN_EMAILS.indexOf(normalized) !== -1;
}

/** Token from `#/admin-edit/{token}` in the page URL hash. */
export function getAdminDeepLinkToken(): string | null {
  const hash = window.location.hash || '';
  const match = hash.match(/#\/admin-edit\/([^/?]+)/i);
  return match?.[1] ?? null;
}

export function isHrDeepLinkMode(): boolean {
  const params = new URLSearchParams(window.location.search);
  return params.get('mode') === 'hr';
}

/** Server-relative path of the page hosting the web part (e.g. /sites/X/SitePages/Forms.aspx). */
export function getFormPageServerRelativePath(): string {
  return `${window.location.pathname}${window.location.search}`;
}

/** Build HR deep link using the current SharePoint page (avoids hard-coded missing .aspx pages). */
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
 * Fix legacy AdminDeepLink values that pointed at a non-existent PFForms.aspx page.
 */
export function repairAdminDeepLinkIfNeeded(
  storedLink: string | undefined,
  token: string
): string {
  if (!storedLink) {
    return buildAdminDeepLinkFromCurrentPage(token);
  }

  if (storedLink.includes('/SitePages/PFForms.aspx')) {
    return buildAdminDeepLinkFromCurrentPage(token);
  }

  return storedLink;
}
