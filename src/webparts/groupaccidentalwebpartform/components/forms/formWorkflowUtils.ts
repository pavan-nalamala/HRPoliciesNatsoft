export type FormKey =
  | 'joiningFormalities'
  | 'teamLifeInsuranceNomination'
  | 'gratuityNominationForm'
  | 'insuranceNominationForm'
  | 'pfDeclaration'
  | 'natItServicesHrPolicyManual'
  | 'natItServicesHandbook';

export const ALL_EMPLOYEE_FORM_KEYS: FormKey[] = [
  'natItServicesHandbook',
  'natItServicesHrPolicyManual',
  'joiningFormalities',
  'teamLifeInsuranceNomination',
  'gratuityNominationForm',
  'insuranceNominationForm',
  'pfDeclaration'
];

export const ADMIN_REVIEW_FORM_KEYS: FormKey[] = [
  'joiningFormalities',
  'teamLifeInsuranceNomination',
  'gratuityNominationForm',
  'insuranceNominationForm',
  'pfDeclaration'
];

export function areAllEmployeeFormsSaved(
  completedForms: Partial<Record<FormKey, boolean>>
): boolean {
  return ALL_EMPLOYEE_FORM_KEYS.every((key) => !!completedForms[key]);
}

export function arePriorPfFormsSaved(
  completedForms: Partial<Record<FormKey, boolean>>
): boolean {
  const prior = ALL_EMPLOYEE_FORM_KEYS.filter((key) => key !== 'pfDeclaration');
  return prior.every((key) => !!completedForms[key]);
}

export function getEmployeeWorkflowStatus(
  completedForms: Partial<Record<FormKey, boolean>>,
  isFinallySubmitted: boolean
): 'Pending' | 'Completed' {
  if (isFinallySubmitted && areAllEmployeeFormsSaved(completedForms)) {
    return 'Completed';
  }
  return 'Pending';
}

export function getAdminWorkflowStatus(hrStatus?: string): 'Pending' | 'Completed' {
  return hrStatus === 'Completed' ? 'Completed' : 'Pending';
}

export function getSubmitButtonLabel(formKey: FormKey): 'Continue' | 'Submit' {
  return formKey === 'pfDeclaration' ? 'Submit' : 'Continue';
}
