import * as React from 'react';
import FormSection from './FormSection';
import { Form, Row, Col, Table, Card, Button } from 'react-bootstrap';
import { useFormik } from 'formik';
import * as Yup from 'yup';
import type { ISequentialFormProps } from './ISequentialFormProps';
import {
  createDateValueChangeHandler,
  createMatchingDateValidation,
  createDateValidation,
  DatePickerInput
} from './dateFieldUtils';
import SignatureUpload from './SignatureUpload';
import moment from 'moment';
import { getSP } from '../../../../pnpjsConfig';
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import {
  candidateFieldNames,
  getCandidateDateValue,
  getCandidateValue,
  setFieldIfEmpty
} from './candidateAutoFillUtils';

type NomineeRow = {
  nomineeNameAndAddress: string;
  relationship: string;
  dateOfBirth: string;
  shareAmount: string;
  guardianDetails: string;
};

type TeamLifeInsuranceFormValues = {
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
  signature: string;
};

const createEmptyNominee = (): NomineeRow => ({
  nomineeNameAndAddress: '',
  relationship: '',
  dateOfBirth: '',
  shareAmount: '',
  guardianDetails: '',
});

const TeamLifeInsuranceNomination = ({
  onComplete,
  sharedEmployeeSignature,
  onEmployeeSignatureChange,
  context,
  employeePFData,
  sharedDateOfBirth,
}: ISequentialFormProps): JSX.Element => {

  const formRef = React.useRef<HTMLDivElement>(null);
  const headerStyle = {
    backgroundColor: '#f18200',
    color: '#fff',
    padding: '10px 15px',
    borderRadius: '8px',
    fontWeight: 600
  };

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

  const commonDateOfBirth =
    sharedDateOfBirth || getCandidateDateValue(employeePFData, candidateFieldNames.dateOfBirth);
  const commonDateOfBirthMessage = commonDateOfBirth
    ? `Date of Birth must match the employee Date of Birth (${commonDateOfBirth})`
    : 'Date of Birth must match the employee Date of Birth';

  const validationSchema = Yup.object().shape({
    employeeName: Yup.string().required('Name of the Employee is required'),
    fatherOrHusbandName: Yup.string().required("Father's / Husband's Name is required"),
    dateOfBirth: createMatchingDateValidation(
      createDateValidation('Date of Birth is required', 'Date of Birth must be in DD/MM/YYYY format'),
      commonDateOfBirth,
      commonDateOfBirthMessage
    )
      .test(
        'dob-age-validation',
        'Minimum age must be 18 years',
        function (value) {
          if (!value) return true;
          const [day, month, year] = value.split('/');
          const dob = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
          const today = new Date();
          const age = today.getFullYear() - dob.getFullYear();
          const monthDiff = today.getMonth() - dob.getMonth();
          const dayDiff = today.getDate() - dob.getDate();
          const actualAge = monthDiff < 0 || (monthDiff === 0 && dayDiff < 0) ? age - 1 : age;
          return actualAge >= 18;
        }
      ),
    sex: Yup.string().required('Sex is required'),
    employeeId: Yup.string(),
    address: Yup.string().required('Address is required'),
    declarationEmployeeName: Yup.string().required('Employee name is required'),
    declarationDate: createDateValidation('Declaration date is required', 'Declaration date must be in DD/MM/YYYY format'),
    nominees: Yup.array()
      .of(
        Yup.object().shape({
          nomineeNameAndAddress: Yup.string().required('Nominee name and address is required'),
          relationship: Yup.string().required('Relationship is required'),
          dateOfBirth: createDateValidation('Nominee date of birth is required', 'Nominee date of birth must be in DD/MM/YYYY format')
            .test(
              'nominee-dob-age-validation',
              'Minimum age must be 18 years',
              function (value) {
                if (!value) return true;
                const [day, month, year] = value.split('/');
                const dob = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                const today = new Date();
                const age = today.getFullYear() - dob.getFullYear();
                const monthDiff = today.getMonth() - dob.getMonth();
                const dayDiff = today.getDate() - dob.getDate();
                const actualAge = monthDiff < 0 || (monthDiff === 0 && dayDiff < 0) ? age - 1 : age;
                return actualAge >= 18;
              }
            ),
          shareAmount: Yup.string().required('Share amount is required'),
          guardianDetails: Yup.string().required('Guardian details is required'),
        })
      )
      .min(1, 'At least one nominee is required')
      .test(
        'total-share-100',
        'Share % across nominees must total exactly 100%',
        function (nominees) {
          if (!Array.isArray(nominees) || nominees.length === 0) return true;
          const total = nominees.reduce((sum, nominee) => {
            const share = parseFloat(nominee.shareAmount) || 0;
            return sum + share;
          }, 0);
          return total === 100;
        }
      ),
    place: Yup.string().required('Place is required'),
    date: createDateValidation('Date is required', 'Date must be in DD/MM/YYYY format'),
    signature: Yup.string().required('Signature of Employee is required'),
  });

  const formik = useFormik<TeamLifeInsuranceFormValues>({
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
      signature: '',
    },
    validationSchema,
    onSubmit: async (values) => {
      try {
        const sp = getSP(context);

        const response = await sp.web.lists
          .getByTitle("InsuranceNomination")
          .items.add({
            Title: values.employeeName,
            EmployeeName: values.employeeName,
            FatherName: values.fatherOrHusbandName,
            DOB: moment(values.dateOfBirth, 'DD/MM/YYYY').toISOString(),
            Gender: values.sex,
            EmployeeID: Number(values.employeeId || null),
            Address: values.address,
            DeclarationDate: moment(values.declarationDate, 'DD/MM/YYYY').toISOString(),
            Place: values.place,
            can_id: String(employeePFData?.ID),
          });

        const parentId = response?.data?.Id;

        const itemId = response.data.Id;

        if (values.signature) {
          const base64 = values.signature.split(",")[1];

          const byteCharacters = atob(base64);
          const byteNumbers = new Array(byteCharacters.length);

          for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
          }

          const blob = new Blob([new Uint8Array(byteNumbers)], { type: "image/png" });

          await sp.web.lists
            .getByTitle("InsuranceNomination")
            .items.getById(itemId)
            .attachmentFiles.add("signature.png", blob);
        }

        if (Array.isArray(values.nominees)) {
          for (const nominee of values.nominees) {

            if (!nominee.nomineeNameAndAddress) continue;

            await sp.web.lists.getByTitle("InsuranceNominees").items.add({
              Title: values.employeeName,
              ParentID: parentId,
              NomineeName: nominee.nomineeNameAndAddress,
              Relationship: nominee.relationship,
              DOB: moment(nominee.dateOfBirth, 'DD/MM/YYYY').toISOString(),
              ShareAmount: Number(nominee.shareAmount),
              GuardianDetails: nominee.guardianDetails
            });
          }
        }

        onComplete?.();

      } catch (error) {
        console.error("Submit Error:", error);
        alert("Submission failed");
      }
    }
  });

  React.useEffect(() => {
    const autoFill = async (): Promise<void> => {
      const employeeName = getCandidateValue(employeePFData, candidateFieldNames.employeeName);

      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'employeeName', employeeName);
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'declarationEmployeeName', employeeName);
      await setFieldIfEmpty(
        formik.values,
        formik.setFieldValue,
        'fatherOrHusbandName',
        getCandidateValue(employeePFData, candidateFieldNames.fatherOrHusbandName)
      );
      await setFieldIfEmpty(
        formik.values,
        formik.setFieldValue,
        'dateOfBirth',
        getCandidateDateValue(employeePFData, candidateFieldNames.dateOfBirth)
      );
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'sex', getCandidateValue(employeePFData, candidateFieldNames.gender));
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'employeeId', getCandidateValue(employeePFData, candidateFieldNames.employeeId));
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'address', getCandidateValue(employeePFData, candidateFieldNames.address));
    };

    void autoFill();
  }, [employeePFData]);

  React.useEffect(() => {
    if (sharedEmployeeSignature && formik.values.signature !== sharedEmployeeSignature) {
      formik.setFieldValue('signature', sharedEmployeeSignature).catch(() => undefined);
    }
  }, [sharedEmployeeSignature]);

  const downloadPDF = async () => {
    const input = formRef.current;
    if (!input) return;

    const pdf = new jsPDF("p", "mm", "a4");

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    const margin = 10;

    const canvas = await html2canvas(input, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
      scrollY: 0,
      onclone: (clonedDocument) => {
        clonedDocument
          .querySelectorAll<HTMLElement>('[data-pdf-hide="true"]')
          .forEach((element) => element.remove());
      }
    });

    const imgWidth = pageWidth - margin * 2;
    const pageHeightContent = pageHeight - margin * 2;
    const pageCanvasHeight = Math.floor((pageHeightContent * canvas.width) / imgWidth);
    const inputRect = input.getBoundingClientRect();
    const canvasScale = canvas.width / inputRect.width;
    const keepTogetherRanges = Array.from(input.querySelectorAll<HTMLElement>('[data-pdf-keep-together="true"]'))
      .map((element) => {
        const rect = element.getBoundingClientRect();
        const start = Math.max(0, Math.floor((rect.top - inputRect.top) * canvasScale));
        const end = Math.min(canvas.height, Math.ceil((rect.bottom - inputRect.top) * canvasScale));

        return { start, end };
      })
      .filter((range) => range.end > range.start && range.end - range.start < pageCanvasHeight);

    let sourceY = 0;

    while (sourceY < canvas.height) {
      let currentPageCanvasHeight = Math.min(pageCanvasHeight, canvas.height - sourceY);
      const pageEnd = sourceY + currentPageCanvasHeight;
      const splitRange = keepTogetherRanges.find(
        (range) => range.start > sourceY && range.start < pageEnd && range.end > pageEnd
      );

      if (splitRange) {
        currentPageCanvasHeight = splitRange.start - sourceY;
      }

      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = currentPageCanvasHeight;

      const pageContext = pageCanvas.getContext("2d");
      if (!pageContext) return;

      pageContext.drawImage(
        canvas,
        0,
        sourceY,
        canvas.width,
        currentPageCanvasHeight,
        0,
        0,
        canvas.width,
        currentPageCanvasHeight
      );

      if (sourceY > 0) pdf.addPage();

      const pageImgData = pageCanvas.toDataURL("image/png");
      const pageImgHeight = (currentPageCanvasHeight * imgWidth) / canvas.width;

      pdf.addImage(
        pageImgData,
        "PNG",
        margin,
        margin,
        imgWidth,
        pageImgHeight
      );

      sourceY += currentPageCanvasHeight;
    }

    pdf.save("Team-Life-Insurance-Nomination.pdf");
  };

  return (
    <>
      <FormSection title="TeamLifeInsuranceNomination" />
      <div ref={formRef} className="no-break">

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
              <h5
                className="text-white p-2 rounded mb-3"
                style={{ backgroundColor: '#f18200' }}
              >
                Employee Details
              </h5>

              <Form.Group className="mb-3">
                <Form.Label>1. Name of the Employee</Form.Label>
                <Form.Control
                  name="employeeName"
                  type="text"
                  placeholder="Enter Full Name"
                  value={formik.values.employeeName}
                  onChange={formik.handleChange}
                  onBlur={formik.handleBlur}
                />
                {formik.touched.employeeName && formik.errors.employeeName && (
                  <p className="text-danger small mb-0">{formik.errors.employeeName}</p>
                )}
              </Form.Group>

              <Form.Group className="mb-3">
                <Form.Label>2. Father's / Husband's Name</Form.Label>
                <Form.Control
                  name="fatherOrHusbandName"
                  type="text"
                  placeholder="Enter Name"
                  value={formik.values.fatherOrHusbandName}
                  onChange={formik.handleChange}
                  onBlur={formik.handleBlur}
                />
                {formik.touched.fatherOrHusbandName && formik.errors.fatherOrHusbandName && (
                  <p className="text-danger small mb-0">{formik.errors.fatherOrHusbandName}</p>
                )}
              </Form.Group>

              <Row className="g-3 mb-3">
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
                    <option>Others</option>
                  </Form.Select>
                  {formik.touched.sex && formik.errors.sex && (
                    <p className="text-danger small mb-0">{formik.errors.sex}</p>
                  )}
                </Col>

                {employeePFData?.EmailID === 'hr@natit.in' && (
                  <Col md={4}>
                    <Form.Label>5. EMP ID</Form.Label>
                    <Form.Control
                      name="employeeId"
                      type="text"
                      placeholder="Enter ID"
                      value={formik.values.employeeId}
                      onChange={formik.handleChange}
                      onBlur={formik.handleBlur}
                      disabled={employeePFData?.EmailID !== 'hr@natit.in'}
                    />
                  </Col>
                )}
              </Row>

              <Form.Group className="mb-4">
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
              </Form.Group>

              <h5
                className="text-white p-2 rounded mb-3"
                style={{ backgroundColor: '#f18200' }}
              >
                Declaration
              </h5>

              <div className="border rounded p-3 bg-light mb-4">
                <div className="d-flex flex-wrap align-items-start gap-2">
                  <span className="mt-2">I,</span>
                  <div>
                    <Form.Control
                      name="declarationEmployeeName"
                      type="text"
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

                  <span className="mt-2">am employed with the above organization since</span>

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
                  I hereby nominate the person(s) mentioned below who is/are member(s) of my
                  family, and confer on him/them the right to receive, to the extent specified
                  below any amount that may be sanctioned by the NAT IT Services Pvt Ltd Group
                  Term Life Insurance in the event of my death while in service in NAT IT
                  Services Pvt Ltd.
                </p>
              </div>



              <div className="d-flex justify-content-between align-items-center">
                <h5 style={headerStyle} className="mb-0">Nominee Details</h5>
                {formik.values.nominees.length < 4 && (
                  <div>
                    <Button
                      type="button"
                      className="border-0"
                      style={{ backgroundColor: '#f18200' }}
                      disabled={
                        !formik.values.nominees[
                          formik.values.nominees.length - 1
                        ]?.nomineeNameAndAddress?.trim() ||
                        !formik.values.nominees[
                          formik.values.nominees.length - 1
                        ]?.relationship?.trim() ||
                        !formik.values.nominees[
                          formik.values.nominees.length - 1
                        ]?.nomineeNameAndAddress
                      }
                      onClick={() => {
                        formik
                          .setFieldValue('nominees', [
                            ...formik.values.nominees,
                            createEmptyNominee(),
                          ])
                          .catch(() => undefined);
                      }}
                    >
                      Add One More
                    </Button>


                    {(!formik.values.nominees[
                      formik.values.nominees.length - 1
                    ]?.nomineeNameAndAddress?.trim() ||
                      !formik.values.nominees[
                        formik.values.nominees.length - 1
                      ]?.relationship?.trim() ||
                      !formik.values.nominees[
                        formik.values.nominees.length - 1
                      ]?.shareAmount) && (
                        <p className="text-danger small mt-1 mb-0">
                          Please fill current nominee details before adding another nominee.
                        </p>
                      )}
                  </div>
                )}
              </div>


              <Table bordered responsive className="align-middle text-center">
                <thead className="table-light">
                  <tr>
                    <th>Name of Nominee(s) & Address</th>
                    <th>Nominee's relationship with the Employee</th>
                    <th style={{ whiteSpace: 'nowrap', minWidth: '260px' }}>Date of Birth</th>
                    <th>Total amount or share of the insured amount to be paid to each nominee</th>
                    <th>
                      If the nominee is minor name and address of the guardian who may receive
                      the amount during the minority of the nominee
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {formik.values.nominees.map((nominee, index) => {
                    const touchedNominee = formik.touched.nominees?.[index];
                    const errorNominee = formik.errors.nominees?.[index];

                    return (
                      <tr key={index}>
                        <td style={{ minWidth: '250px' }}>
                          <Form.Control
                            name={`nominees[${index}].nomineeNameAndAddress`}
                            placeholder="Name and Address"
                            value={nominee.nomineeNameAndAddress}
                            onChange={formik.handleChange}
                            onBlur={formik.handleBlur}
                            style={{ width: '100%' }}
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

                        <td style={{ minWidth: '260px', whiteSpace: 'nowrap' }}>
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
                            onChange={(e) => {
                              let value = e.target.value.replace(/[^0-9]/g, '');

                              // Prevent above 100 for single field
                              if (Number(value) > 100) {
                                value = '100';
                              }

                              // Calculate total percentage
                              const updatedNominees = [...formik.values.nominees];
                              updatedNominees[index].shareAmount = value;

                              const total = updatedNominees.reduce(
                                (sum, item) => sum + Number(item.shareAmount || 0),
                                0
                              );

                              // Allow update only if total <= 100
                              if (total <= 100) {
                                formik
                                  .setFieldValue(
                                    `nominees[${index}].shareAmount`,
                                    value
                                  )
                                  .catch(() => undefined);
                              } else {
                                formik.setFieldError(
                                  `nominees[${index}].shareAmount`,
                                  'Total share amount cannot exceed 100%'
                                );
                              }
                            }}
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
                          <div className="d-flex align-items-start gap-2">

                            <div style={{ flex: 1 }}>
                              <Form.Control
                                as="textarea"
                                rows={1}
                                name={`nominees[${index}].guardianDetails`}
                                placeholder="Guardian Details"
                                value={nominee.guardianDetails}
                                onChange={(e) => {
                                  formik.handleChange(e);

                                  e.target.style.height = "auto";
                                  e.target.style.height =
                                    e.target.scrollHeight + "px";
                                }}
                                onBlur={formik.handleBlur}
                                style={{
                                  overflow: "hidden",
                                  resize: "none"
                                }}
                              />

                              {typeof touchedNominee === "object" &&
                                touchedNominee?.guardianDetails &&
                                typeof errorNominee === "object" &&
                                errorNominee?.guardianDetails && (
                                  <p className="text-danger small mb-0 text-start">
                                    {errorNominee.guardianDetails}
                                  </p>
                                )}
                            </div>

                            {formik.values.nominees.length > 1 && (
                              <Button
                                type="button"
                                variant="danger"
                                size="sm"
                                onClick={() => {

                                  const updatedNominees =
                                    formik.values.nominees.filter(
                                      (_, i) => i !== index
                                    );

                                  formik
                                    .setFieldValue('nominees', updatedNominees)
                                    .catch(() => undefined);
                                }}
                              >
                                Remove
                              </Button>
                            )}

                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </Table>

              {formik.errors.nominees && typeof formik.errors.nominees === 'string' && (
                <div className="alert alert-danger mb-3" role="alert">
                  {formik.errors.nominees}
                </div>
              )}

              {formik.values.nominees.reduce(
                (sum, n) => sum + (parseFloat(n.shareAmount) || 0),
                0
              ) > 100 && (
                  <div className="mb-3 p-2 bg-light border rounded">
                    <p className="mb-0 fw-semibold">
                      <div className="text-danger">
                        Share should not exceed 100%
                      </div>

                    </p>
                  </div>
                )}

              <Row className="mt-4 g-3">
                <Col md={6}>
                  <Form.Group className="mb-3">
                    <Form.Label>Place</Form.Label>
                    <Form.Control
                      name="place"
                      type="text"
                      value={formik.values.place}
                      onChange={formik.handleChange}
                      onBlur={formik.handleBlur}
                    />
                    {formik.touched.place && formik.errors.place && (
                      <p className="text-danger small mb-0">{formik.errors.place}</p>
                    )}
                  </Form.Group>

                  <Form.Group>
                    <Form.Label>Date</Form.Label>
                    <DatePickerInput
                      name="date"
                      value={formik.values.date}
                      onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'date')}
                      onBlur={formik.handleBlur}
                    />
                    {formik.touched.date && formik.errors.date && (
                      <p className="text-danger small mb-0">{formik.errors.date}</p>
                    )}
                  </Form.Group>
                </Col>

                <Col md={6} className="d-flex flex-column justify-content-end">
                  <Form.Label>Signature of Employee</Form.Label>
                  <SignatureUpload
                    name="signature"
                    value={formik.values.signature}
                    onChange={(value) => {
                      formik.setFieldValue('signature', value).catch(() => undefined);
                      onEmployeeSignatureChange?.(value);
                    }}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.signature && formik.errors.signature && (
                    <p className="text-danger small mb-0">{formik.errors.signature}</p>
                  )}
                </Col>
              </Row>

              <div className="text-center mt-4">
                <Button
                  type="submit"
                  className="border-0"
                  style={{ backgroundColor: '#f18200' }}
                  disabled={formik.isSubmitting}
                >
                  {formik.isSubmitting ? 'Submitting...' : 'Submit Form'}
                </Button>
                {employeePFData?.EmailID === 'hr@natit.in' && (
                  <Button
                    type="button"
                    className="border-0 ms-2"
                    style={{ backgroundColor: "#f18200" }}
                    onClick={downloadPDF}
                  >
                    Download PDF
                  </Button>
                )}

              </div>
            </Form>
          </Card.Body>
        </Card>
      </div>
    </>
  );
};

export default TeamLifeInsuranceNomination;
