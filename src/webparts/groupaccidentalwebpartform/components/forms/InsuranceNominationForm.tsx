import * as React from 'react';
import FormSection from './FormSection';
import {
  Card,
  Form,
  Row,
  Col,
  Table,
  Button
} from 'react-bootstrap';
import { useFormik } from 'formik';
import * as Yup from 'yup';
import type { ISequentialFormProps } from './ISequentialFormProps';
import {
  createDateValueChangeHandler,
  createDateValidation,
  DatePickerInput
} from './dateFieldUtils';
import SignatureUpload from './SignatureUpload';

type NomineeRow = {
  nomineeNameAndAddress: string;
  relationship: string;
  dateOfBirth: string;
  shareAmount: string;
  guardianDetails: string;
};

type InsuranceNominationValues = {
  employeeName: string;
  fatherOrHusbandName: string;
  dateOfBirth: string;
  sex: string;
  employeeId: string;
  address: string;
  declarationEmployeeName: string;
  declarationDate: string;
  nominees: NomineeRow[];
  place: string;
  date: string;
  employeeSignature: string;
};

const createEmptyNominee = (): NomineeRow => ({
  nomineeNameAndAddress: '',
  relationship: '',
  dateOfBirth: '',
  shareAmount: '',
  guardianDetails: '',
});

const InsuranceNominationForm = ({
  onComplete,
  sharedEmployeeSignature,
  onEmployeeSignatureChange
}: ISequentialFormProps): JSX.Element => {
  React.useEffect(() => {
    if (!document.getElementById('bootstrap-css')) {
      const link = document.createElement('link');
      link.id = 'bootstrap-css';
      link.rel = 'stylesheet';
      link.href =
        'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css';
      document.head.appendChild(link);
    }
  }, []);

  const headerStyle = {
    backgroundColor: '#f18200',
    color: '#fff',
    padding: '10px 15px',
    borderRadius: '8px',
    fontWeight: 600
  };

  const validationSchema = Yup.object().shape({
    employeeName: Yup.string().required('Employee Name is required'),
    fatherOrHusbandName: Yup.string().required('Father / Husband Name is required'),
    dateOfBirth: createDateValidation('Date of Birth is required', 'Date of Birth must be in DD/MM/YYYY format'),
    sex: Yup.string().required('Sex is required'),
    employeeId: Yup.string().required('EMP ID is required'),
    address: Yup.string().required('Address is required'),
    declarationEmployeeName: Yup.string().required('Declaration employee name is required'),
    declarationDate: createDateValidation('Declaration date is required', 'Declaration date must be in DD/MM/YYYY format'),
    nominees: Yup.array()
      .of(
        Yup.object().shape({
          nomineeNameAndAddress: Yup.string().required('Nominee Name & Address is required'),
          relationship: Yup.string().required('Relationship is required'),
          dateOfBirth: createDateValidation('Nominee Date of Birth is required', 'Nominee Date of Birth must be in DD/MM/YYYY format'),
          shareAmount: Yup.string()
            .required('Share Amount is required')
            .matches(/^\d+(\.\d+)?$/, 'Share Amount must be a valid number'),
          guardianDetails: Yup.string().required('Guardian Details is required'),
        })
      )
      .min(1, 'At least one nominee is required'),
    place: Yup.string().required('Place is required'),
    date: createDateValidation('Date is required', 'Date must be in DD/MM/YYYY format'),
    employeeSignature: Yup.string().required('Employee Signature is required'),
  });

  const formik = useFormik<InsuranceNominationValues>({
    initialValues: {
      employeeName: '',
      fatherOrHusbandName: '',
      dateOfBirth: '',
      sex: '',
      employeeId: '',
      address: '',
      declarationEmployeeName: '',
      declarationDate: '',
      nominees: [createEmptyNominee()],
      place: '',
      date: '',
      employeeSignature: '',
    },
    validationSchema,
    onSubmit: async (values) => {
      console.log('values', values);
      onComplete?.();
    },
  });

  React.useEffect(() => {
    if (sharedEmployeeSignature && formik.values.employeeSignature !== sharedEmployeeSignature) {
      formik.setFieldValue('employeeSignature', sharedEmployeeSignature).catch(() => undefined);
    }
  }, [sharedEmployeeSignature]);

  return (
    <>
      <FormSection title="Insurance Nomination Form" />

      <div>
        <Card className="shadow border-0 rounded-4 mx-auto">
          <Card.Body className="p-4 p-md-5">
            <div className="text-center mb-4">
              <h3
                className="fw-bold text-white py-2 rounded"
                style={{ backgroundColor: '#f18200' }}
              >
                NOMINATION FORM FOR TERM LIFE INSURANCE
              </h3>
            </div>

            <div className="mb-4">
              <p className="mb-1 fw-bold">TO</p>
              <p className="mb-1 fw-semibold">NAT IT Services Pvt Ltd</p>
              <p className="mb-0">
                Plot No 21, Sruthi Sadan, Serlingampalli Mandal, Gachibowli, Hyderabad -
                500032
              </p>
            </div>

            <Form onSubmit={formik.handleSubmit}>
              <h5 style={headerStyle}>Employee Details</h5>

              <Row className="g-3 mt-2">
                <Col md={6}>
                  <Form.Label>1. Employee Name</Form.Label>
                  <Form.Control
                    name="employeeName"
                    placeholder="Enter Full Name"
                    value={formik.values.employeeName}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.employeeName && formik.errors.employeeName && (
                    <p className="text-danger small mb-0">{formik.errors.employeeName}</p>
                  )}
                </Col>

                <Col md={6}>
                  <Form.Label>2. Father / Husband Name</Form.Label>
                  <Form.Control
                    name="fatherOrHusbandName"
                    placeholder="Enter Name"
                    value={formik.values.fatherOrHusbandName}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.fatherOrHusbandName && formik.errors.fatherOrHusbandName && (
                    <p className="text-danger small mb-0">{formik.errors.fatherOrHusbandName}</p>
                  )}
                </Col>

                <Col md={4}>
                  <Form.Label>3. Date of Birth</Form.Label>
                  <DatePickerInput
                    name="dateOfBirth"
                    value={formik.values.dateOfBirth}
                    onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'dateOfBirth')}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.dateOfBirth && formik.errors.dateOfBirth && (
                    <p className="text-danger small mb-0">{formik.errors.dateOfBirth}</p>
                  )}
                </Col>

                <Col md={4}>
                  <Form.Label>4. Sex</Form.Label>
                  <Form.Select
                    name="sex"
                    value={formik.values.sex}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  >
                    <option value="">Select</option>
                    <option>Male</option>
                    <option>Female</option>
                    <option>Other</option>
                  </Form.Select>
                  {formik.touched.sex && formik.errors.sex && (
                    <p className="text-danger small mb-0">{formik.errors.sex}</p>
                  )}
                </Col>

                <Col md={4}>
                  <Form.Label>5. EMP ID</Form.Label>
                  <Form.Control
                    name="employeeId"
                    placeholder="Enter ID"
                    value={formik.values.employeeId}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.employeeId && formik.errors.employeeId && (
                    <p className="text-danger small mb-0">{formik.errors.employeeId}</p>
                  )}
                </Col>

                <Col md={12}>
                  <Form.Label>6. Address</Form.Label>
                  <Form.Control
                    name="address"
                    as="textarea"
                    rows={3}
                    placeholder="Enter Address"
                    value={formik.values.address}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.address && formik.errors.address && (
                    <p className="text-danger small mb-0">{formik.errors.address}</p>
                  )}
                </Col>
              </Row>

              <h5 style={headerStyle} className="mt-4">
                Declaration
              </h5>

              <Card className="border-0 bg-light mt-3">
                <Card.Body>
                  <div className="d-flex flex-wrap gap-2 align-items-start">
                    <span className="mt-2">I,</span>

                    <div>
                      <Form.Control
                        name="declarationEmployeeName"
                        style={{ width: '220px' }}
                        placeholder="Employee Name"
                        value={formik.values.declarationEmployeeName}
                        onChange={formik.handleChange}
                        onBlur={formik.handleBlur}
                      />
                      {formik.touched.declarationEmployeeName &&
                        formik.errors.declarationEmployeeName && (
                          <p className="text-danger small mb-0">
                            {formik.errors.declarationEmployeeName}
                          </p>
                        )}
                    </div>

                    <span className="mt-2">
                      am employed with the above organization since
                    </span>

                    <div>
                      <div style={{ width: '220px' }}>
                        <DatePickerInput
                          name="declarationDate"
                          value={formik.values.declarationDate}
                          onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'declarationDate')}
                          onBlur={formik.handleBlur}
                        />
                      </div>
                      {formik.touched.declarationDate && formik.errors.declarationDate && (
                        <p className="text-danger small mb-0">{formik.errors.declarationDate}</p>
                      )}
                    </div>
                  </div>

                  <p className="mt-3 mb-0">
                    I hereby nominate the person(s) mentioned below who is/are member(s) of
                    my family, and confer on them the right to receive the amount payable
                    under the policy.
                  </p>
                </Card.Body>
              </Card>

              <div className="d-flex justify-content-between align-items-center mt-4">
                <h5 style={headerStyle} className="mb-0">
                  Nominee Details
                </h5>
                <Button
                  type="button"
                  className="border-0"
                  style={{ backgroundColor: '#f18200' }}
                  onClick={() => {
                    formik
                      .setFieldValue('nominees', [...formik.values.nominees, createEmptyNominee()])
                      .catch(() => undefined);
                  }}
                >
                  Add One More
                </Button>
              </div>

              <Table
                bordered
                responsive
                className="mt-3 text-center align-middle"
              >
                <thead className="table-light">
                  <tr>
                    <th>Name of Nominee(s) & Address</th>
                    <th>Relationship</th>
                    <th>Date of Birth</th>
                    <th>Share Amount</th>
                    <th>Guardian Details (If Minor)</th>
                  </tr>
                </thead>

                <tbody>
                  {formik.values.nominees.map((nominee, index) => {
                    const touchedNominee = formik.touched.nominees?.[index];
                    const errorNominee = formik.errors.nominees?.[index];

                    return (
                      <tr key={index}>
                        <td>
                          <Form.Control
                            name={`nominees[${index}].nomineeNameAndAddress`}
                            as="textarea"
                            rows={2}
                            placeholder="Enter Name & Address"
                            value={nominee.nomineeNameAndAddress}
                            onChange={formik.handleChange}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedNominee === 'object' &&
                            touchedNominee?.nomineeNameAndAddress &&
                            typeof errorNominee === 'object' &&
                            errorNominee?.nomineeNameAndAddress && (
                              <p className="text-danger small mb-0 text-start">
                                {errorNominee.nomineeNameAndAddress}
                              </p>
                            )}
                        </td>

                        <td>
                          <Form.Control
                            name={`nominees[${index}].relationship`}
                            placeholder="Relationship"
                            value={nominee.relationship}
                            onChange={formik.handleChange}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedNominee === 'object' &&
                            touchedNominee?.relationship &&
                            typeof errorNominee === 'object' &&
                            errorNominee?.relationship && (
                              <p className="text-danger small mb-0 text-start">
                                {errorNominee.relationship}
                              </p>
                            )}
                        </td>

                        <td>
                          <DatePickerInput
                            name={`nominees[${index}].dateOfBirth`}
                            value={nominee.dateOfBirth}
                            onValueChange={createDateValueChangeHandler(formik.setFieldValue, `nominees[${index}].dateOfBirth`)}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedNominee === 'object' &&
                            touchedNominee?.dateOfBirth &&
                            typeof errorNominee === 'object' &&
                            errorNominee?.dateOfBirth && (
                              <p className="text-danger small mb-0 text-start">
                                {errorNominee.dateOfBirth}
                              </p>
                            )}
                        </td>

                        <td>
                          <Form.Control
                            name={`nominees[${index}].shareAmount`}
                            placeholder="%"
                            value={nominee.shareAmount}
                            onChange={formik.handleChange}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedNominee === 'object' &&
                            touchedNominee?.shareAmount &&
                            typeof errorNominee === 'object' &&
                            errorNominee?.shareAmount && (
                              <p className="text-danger small mb-0 text-start">
                                {errorNominee.shareAmount}
                              </p>
                            )}
                        </td>

                        <td>
                          <Form.Control
                            name={`nominees[${index}].guardianDetails`}
                            as="textarea"
                            rows={2}
                            placeholder="Guardian Details"
                            value={nominee.guardianDetails}
                            onChange={formik.handleChange}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedNominee === 'object' &&
                            touchedNominee?.guardianDetails &&
                            typeof errorNominee === 'object' &&
                            errorNominee?.guardianDetails && (
                              <p className="text-danger small mb-0 text-start">
                                {errorNominee.guardianDetails}
                              </p>
                            )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </Table>

              <Row className="g-3 mt-4">
                <Col md={6}>
                  <Form.Label>Place</Form.Label>
                  <Form.Control
                    name="place"
                    placeholder="Enter Place"
                    value={formik.values.place}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.place && formik.errors.place && (
                    <p className="text-danger small mb-0">{formik.errors.place}</p>
                  )}

                  <Form.Label className="mt-3">Date</Form.Label>
                  <DatePickerInput
                    name="date"
                    value={formik.values.date}
                    onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'date')}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.date && formik.errors.date && (
                    <p className="text-danger small mb-0">{formik.errors.date}</p>
                  )}
                </Col>

                <Col md={6} className="d-flex flex-column justify-content-end">
                  <Form.Label>Employee Signature</Form.Label>
                  <SignatureUpload
                    name="employeeSignature"
                    value={formik.values.employeeSignature}
                    onChange={(value) => {
                      formik.setFieldValue('employeeSignature', value).catch(() => undefined);
                      onEmployeeSignatureChange?.(value);
                    }}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.employeeSignature && formik.errors.employeeSignature && (
                    <p className="text-danger small mb-0">{formik.errors.employeeSignature}</p>
                  )}
                </Col>
              </Row>

              <div className="text-center mt-5">
                <Button
                  type="submit"
                  size="lg"
                  style={{
                    backgroundColor: '#f18200',
                    border: 'none',
                    minWidth: '220px'
                  }}
                >
                  Submit Form
                </Button>
              </div>
            </Form>
          </Card.Body>
        </Card>
      </div>
    </>
  );
};

export default InsuranceNominationForm;
