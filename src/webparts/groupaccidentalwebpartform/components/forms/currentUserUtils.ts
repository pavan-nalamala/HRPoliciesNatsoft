import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../../pnpjsConfig';
import { normalizeUserEmail } from './hrAccessUtils';
import '@pnp/sp/profiles';

export type CurrentUserProfile = {
  displayName: string;
  mail: string;
  userPrincipalName: string;
};

type UserProfileProperty = {
  Key: string;
  Value: string;
};

function getProfileProperty(
  properties: UserProfileProperty[] | undefined,
  key: string
): string {
  const match = properties?.find((item) => item.Key === key);
  return match?.Value?.trim() ?? '';
}

function emailFromAccountName(accountName: string): string {
  return normalizeUserEmail(accountName);
}

type SharePointUserProfile = {
  DisplayName?: string;
  UserProfileProperties?: UserProfileProperty[];
};

/**
 * Load the current user via PnPjs (SharePoint user profile), not Microsoft Graph.
 *
 * const sp = getSP(context);
 * const profile = await sp.profiles.myProperties();
 */
export async function getCurrentUserProfile(
  context: WebPartContext
): Promise<CurrentUserProfile> {
  const sp = getSP(context);
  const profile = (await sp.profiles.myProperties()) as SharePointUserProfile;

  const properties = profile.UserProfileProperties;

  const workEmail = getProfileProperty(properties, 'WorkEmail');
  const spsEmail = getProfileProperty(properties, 'SPS-Email');
  const userName = getProfileProperty(properties, 'UserName');
  const accountName = getProfileProperty(properties, 'AccountName');

  const rawMail = workEmail || spsEmail || userName || emailFromAccountName(accountName);
  const rawUpn = userName || emailFromAccountName(accountName) || rawMail;
  const mail = normalizeUserEmail(rawMail);
  const userPrincipalName = normalizeUserEmail(rawUpn) || mail;

  return {
    displayName: profile.DisplayName?.trim() ?? '',
    mail,
    userPrincipalName
  };
}

export function getEmailFromCurrentUser(
  user: CurrentUserProfile | null | undefined
): string {
  if (!user) {
    return '';
  }

  return normalizeUserEmail(user.mail || user.userPrincipalName);
}
