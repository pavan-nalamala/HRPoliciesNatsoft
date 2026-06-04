import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IGroupaccidentalwebpartformProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  context: WebPartContext;
  /** Title of the linked document library (default: Link Document Library). */
  hrDocumentLibraryTitle?: string;
  /** Server-relative path for HR deep links (AdminDeepLink / emails). */
  hrFormPagePath?: string;
}
