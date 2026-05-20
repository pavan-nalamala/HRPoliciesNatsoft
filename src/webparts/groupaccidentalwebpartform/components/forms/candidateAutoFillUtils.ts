import moment from 'moment';

export const getCandidateValue = (
  candidate: any,
  fieldNames: string[]
): string => {
  if (!candidate) return '';

  for (const fieldName of fieldNames) {
    const value = candidate[fieldName];

    if (value !== undefined && value !== null && String(value).trim() !== '') {
      return String(value);
    }
  }

  return '';
};

export const getCandidateDateValue = (
  candidate: any,
  fieldNames: string[]
): string => {
  const value = getCandidateValue(candidate, fieldNames);

  if (!value) return '';

  const date = moment(value);

  return date.isValid() ? date.format('DD/MM/YYYY') : value;
};

export const setFieldIfEmpty = async (
  values: Record<string, any>,
  setFieldValue: (field: string, value: any) => Promise<any>,
  fieldName: string,
  value: any
): Promise<void> => {
  if (value === undefined || value === null || String(value).trim() === '') return;
  if (values[fieldName] !== undefined && String(values[fieldName]).trim() !== '') return;

  await setFieldValue(fieldName, value);
};

export const candidateFieldNames = {
  employeeName: ['EmployeeName', 'FullName', 'CandidateName', 'Name', 'Title'],
  fatherOrHusbandName: ['FatherName', 'FatherOrHusbandName', 'FatherSpouseName', 'Father_x0020_Name'],
  dateOfBirth: ['DateOfBirth', 'DOB', 'Date_x0020_of_x0020_Birth', 'employee_DOB'],
  gender: ['Gender', 'Sex'],
  employeeId: ['EmployeeID', 'EmployeeId', 'EmpID', 'Employee_x0020_ID'],
  address: ['Address', 'PermanentAddress', 'PresentAddress', 'CurrentAddress'],
  department: ['Department', 'DepartmentName', 'Department_x0020_Name'],
  designation: ['Designation', 'Designations', 'JobTitle'],
  joiningDate: ['DateOfJoining', 'JoiningDate', 'DOJ'],
  panNo: ['PANNo', 'PAN', 'PanNo'],
  maritalStatus: ['MaritalStatus'],
  spouseName: ['SpouseName'],
  email: ['Email', 'EmailID', 'Title'],
  mobile: ['Mobile', 'MobileNo', 'Phone', 'ContactNo']
};
