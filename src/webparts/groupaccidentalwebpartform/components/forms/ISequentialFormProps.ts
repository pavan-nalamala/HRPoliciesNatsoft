export interface ISequentialFormProps {
  onComplete?: () => void;
  sharedEmployeeSignature?: string;
  onEmployeeSignatureChange?: (value: string) => void;
}
