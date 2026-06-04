import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISequentialFormProps {
  onComplete?: () => void;

  // Signature sharing
  sharedEmployeeSignature?: string;
  onEmployeeSignatureChange?: (value: string) => void;
  sharedDateOfBirth?: string;
  onSharedDateOfBirthChange?: (value: string) => void;

  // SharePoint context
  spHttpClient?: SPHttpClient;
  siteUrl?: string;
  context?: WebPartContext;
  currentUser?: any;
  employeePFData: any;
  /** Saved progress exists in SharePoint for this tab (employee may still edit). */
  hasSavedProgress?: boolean;
  /** Employee completed final PF submit — all tabs read-only for employee. */
  isFinallySubmitted?: boolean;
  candidateId?: number;
  isAdminEdit?: boolean;
  submitButtonLabel?: 'Continue' | 'Submit';
  workflowStatus?: 'inProgress' | 'pendingHr' | 'verified';
  /** Required before employee can run final PF Submit. */
  canSubmitFinal?: boolean;
  /** SharePoint document library for HR-completed file archive. */
  documentLibraryTitle?: string;
  /** Server-relative form page path for AdminDeepLink (HR emails). */
  hrFormPagePath?: string;
  /** Human-readable HR status label for banners. */
  hrStatusLabel?: string;
}
