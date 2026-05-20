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
  employeePFData:any;
  isSubmitted?: boolean;

}
