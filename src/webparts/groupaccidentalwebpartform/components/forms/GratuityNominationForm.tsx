import * as React from 'react';
import FormSection from './FormSection';
import { Form, Row, Col, Table, Card, Button } from 'react-bootstrap';
import { useFormik } from 'formik';
import * as Yup from 'yup';
import type { ISequentialFormProps } from './ISequentialFormProps';
import {
  createDateValueChangeHandler,
  createDateValidation,
  DatePickerInput,
  formatSharePointDateForInput,
  parseFormDateForSharePoint
} from './dateFieldUtils';
import SignatureUpload from './SignatureUpload';
import { getSP } from '../../../../pnpjsConfig';
import moment from 'moment';
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import {
  candidateFieldNames,
  getCandidateDateValue,
  getCandidateValue,
  setFieldIfEmpty
} from './candidateAutoFillUtils';
import {
  getListItemsByParentId,
  replaceListItemAttachment,
  toAttachmentPreviewUrl,
  upsertByCanId
} from './sharePointListUtils';


type NomineeRow = {
  fullNameAndAddress: string;
  relationship: string;
  age: string;
  sharePercentage: string;
};

type WitnessRow = {
  nameAndAddress: string;
  signature: File | string;
};

type GratuityNominationValues = {
  employeeIntroName: string;
  spouseExclusionDate: string;
  nominees: NomineeRow[];
  employeeStatementNameAndAddress: string;
  sex: string;
  religion: string;
  maritalStatus: string;
  departmentBranchSection: string;
  employeeId: string;
  dateOfJoining: string;
  permanentAddress: string;
  place: string;
  date: string;
  employeeSignature: File | string;
  witnesses: WitnessRow[];
  referenceNo: string;
  employerCertificateDate: string;
  authorizedSignature: File | string;
  designation: string;
  acknowledgmentDate: string;
  acknowledgmentEmployeeSignature: File | string;
};

const createEmptyNominee = (): NomineeRow => ({
  fullNameAndAddress: '',
  relationship: '',
  age: '',
  sharePercentage: '',
});

const createEmptyWitness = (): WitnessRow => ({
  nameAndAddress: '',
  signature: '',
});

const GratuityNominationForm = ({
  onComplete,
  sharedEmployeeSignature,
  onEmployeeSignatureChange,
  context,
  employeePFData,
  isAdminEdit,
  isFinallySubmitted,
  submitButtonLabel = 'Continue',
}: ISequentialFormProps): JSX.Element => {
  const readOnly = !!isFinallySubmitted && !isAdminEdit;

  const formRef = React.useRef<HTMLDivElement>(null);
  const [signaturePreview, setSignaturePreview] = React.useState<any>({});


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
    employeeIntroName: Yup.string().required('Full Name is required'),
    spouseExclusionDate: createDateValidation('Spouse exclusion date is required', 'Spouse exclusion date must be in DD/MM/YYYY format'),
    nominees: Yup.array()
      .of(
        Yup.object().shape({
          fullNameAndAddress: Yup.string().required('Full Name & Address is required'),
          relationship: Yup.string().required('Relationship is required'),
          age: Yup.string()
            .required('Age is required')
            .matches(/^\d+$/, 'Age must be a valid number'),
          sharePercentage: Yup.string()
            .required('Share % is required')
            .matches(/^\d+(\.\d+)?$/, 'Share % must be a valid number'),
        })
      )
      .min(1, 'At least one nominee is required'),
    employeeStatementNameAndAddress: Yup.string().required(
      'Full Name & Permanent Address is required'
    ),
    sex: Yup.string().required('Sex is required'),
    religion: Yup.string().required('Religion is required'),
    maritalStatus: Yup.string().required('Marital Status is required'),
    departmentBranchSection: Yup.string().required('Department / Branch / Section is required'),
    employeeId: Yup.string(),
    dateOfJoining: createDateValidation('Date of Joining is required', 'Date of Joining must be in DD/MM/YYYY format'),
    permanentAddress: Yup.string().required('Permanent Address is required'),
    place: Yup.string().required('Place is required'),
    date: createDateValidation('Date is required', 'Date must be in DD/MM/YYYY format'),
    employeeSignature: Yup.mixed<File | string>()
      .required('Employee Signature is required'),
    witnesses: Yup.array().of(
      Yup.object().shape({
        nameAndAddress: Yup.string().required('Witness Name & Address is required'),
        signature: Yup.mixed<File | string>()
          .required('Witness Signature is required'),
      })
    ),
    referenceNo: Yup.string().required('Reference No is required'),
    employerCertificateDate: createDateValidation('Employer certificate date is required', 'Employer certificate date must be in DD/MM/YYYY format'),
    authorizedSignature: Yup.string().required('Authorized Signature is required'),
    designation: Yup.string().required('Designation is required'),
    acknowledgmentDate: createDateValidation('Acknowledgment date is required', 'Acknowledgment date must be in DD/MM/YYYY format'),
    acknowledgmentEmployeeSignature: Yup.mixed<File | string>().required(
      'Acknowledgment Employee Signature is required'
    ),
  });

  const formik = useFormik<GratuityNominationValues>({
    initialValues: {
      employeeIntroName: '',
      spouseExclusionDate: '',
      nominees: [createEmptyNominee()],
      employeeStatementNameAndAddress: '',
      sex: '',
      religion: '',
      maritalStatus: '',
      departmentBranchSection: '',
      employeeId: '',
      dateOfJoining: '',
      permanentAddress: '',
      place: '',
      date: '',
      employeeSignature: '',
      witnesses: [createEmptyWitness(), createEmptyWitness()],
      referenceNo: '',
      employerCertificateDate: '',
      authorizedSignature: '',
      designation: '',
      acknowledgmentDate: '',
      acknowledgmentEmployeeSignature: '',
    },
    validationSchema,
    onSubmit: async (values) => {
      try {
        const sp = getSP(context);

        const spouseExclusionDate = parseFormDateForSharePoint(
          values.spouseExclusionDate,
          'datetime'
        );
        const dateOfJoining = parseFormDateForSharePoint(values.dateOfJoining, 'datetime');
        const formDate = parseFormDateForSharePoint(values.date, 'datetime');
        const employerCertificateDate = parseFormDateForSharePoint(
          values.employerCertificateDate,
          'datetime'
        );
        const acknowledgmentDate = parseFormDateForSharePoint(
          values.acknowledgmentDate,
          'datetime'
        );

        if (
          !spouseExclusionDate ||
          !dateOfJoining ||
          !formDate ||
          !employerCertificateDate ||
          !acknowledgmentDate
        ) {
          throw new Error(
            'Invalid date on the form. Please enter all dates in DD/MM/YYYY format.'
          );
        }

        const itemId = await upsertByCanId(
          sp,
          'GratuityNomination',
          String(employeePFData?.ID),
          {
            Title: values.employeeIntroName,
            EmployeeName: values.employeeIntroName,
            SpouseExclusionDate: spouseExclusionDate,
            EmployeeStatName: values.employeeStatementNameAndAddress,
            Sex: values.sex,
            Religion: values.religion,
            MaritalStatus: values.maritalStatus,
            Department: values.departmentBranchSection,
            EmployeeId: Number(values.employeeId),
            DateOfJoining: dateOfJoining,
            PermanentAddress: values.permanentAddress,
            Place: values.place,
            FormDate0: formDate,
            ReferenceNo: String(values.referenceNo),
            EmployerCertificateDate: employerCertificateDate,
            Designation: values.designation,
            AcknowledgmentDate: acknowledgmentDate
          }
        );

        const existingNominees = await sp.web.lists
          .getByTitle('GratuityNominees')
          .items.filter(`ParentID eq ${itemId}`)();
        for (const row of existingNominees) {
          await sp.web.lists.getByTitle('GratuityNominees').items.getById(row.Id).delete();
        }
        const existingWitnesses = await sp.web.lists
          .getByTitle('GratuityWitnesses')
          .items.filter(`ParentID eq ${itemId}`)();
        for (const row of existingWitnesses) {
          await sp.web.lists.getByTitle('GratuityWitnesses').items.getById(row.Id).delete();
        }


        const uploadAttachment = async (
          fileName: string,
          file: File | string
        ): Promise<void> => {
          await replaceListItemAttachment(
            sp,
            'GratuityNomination',
            itemId,
            fileName,
            file
          );
        };

        await uploadAttachment('EmployeeSignature.png', values.employeeSignature);
        await uploadAttachment('Witness1Signature.png', values.witnesses?.[0]?.signature);
        await uploadAttachment('Witness2Signature.png', values.witnesses?.[1]?.signature);
        await uploadAttachment('AuthorizedSignature.png', values.authorizedSignature);
        await uploadAttachment(
          'AcknowledgmentEmployeeSignature.png',
          values.acknowledgmentEmployeeSignature
        );

        // 5. SAVE NOMINEES
        if (values.nominees?.length) {
          for (const n of values.nominees) {
            if (!n.fullNameAndAddress) continue;

            await sp.web.lists
              .getByTitle("GratuityNominees")
              .items.add({
                ParentID: itemId,
                FullNameAndAddress: n.fullNameAndAddress,
                Relationship: n.relationship,
                Age: Number(n.age),
                SharePercentage: Number(n.sharePercentage),
              });
          }
        }

        // 6. SAVE WITNESSES
        if (values.witnesses?.length) {
          for (const w of values.witnesses) {
            if (!w.nameAndAddress) continue;

            await sp.web.lists
              .getByTitle("GratuityWitnesses")
              .items.add({
                ParentID: itemId,
                NameAndAddress: w.nameAndAddress,
              });
          }
        }

        onComplete?.();
      } catch (error) {
        console.error("Submit Error:", error);
        alert(
          error instanceof Error
            ? error.message
            : 'Submission failed. Please check all dates are in DD/MM/YYYY format.'
        );
      }
    }
  });

  React.useEffect(() => {

    const loadExistingData = async (): Promise<void> => {

      try {

        const sp = getSP(context);

        // MAIN ITEM
        const items = await sp.web.lists
          .getByTitle("GratuityNomination")
          .items
          .filter(`can_id eq '${employeePFData?.ID}'`)
          .top(1)
          .orderBy("Id", false)();

        if (!items.length) return;

        const item = items[0];
        const parentId = Number(item.Id);

        const attachments = await sp.web.lists
          .getByTitle("GratuityNomination")
          .items.getById(parentId)
          .attachmentFiles();

        const employeeSignature =
          attachments.find(
            (x: any) =>
              x.FileName === "EmployeeSignature.png"
          );

        const witness1Signature =
          attachments.find(
            (x: any) =>
              x.FileName === "Witness1Signature.png"
          );

        const witness2Signature =
          attachments.find(
            (x: any) =>
              x.FileName === "Witness2Signature.png"
          );

        const authorizedSignature =
          attachments.find(
            (x: any) =>
              x.FileName === "AuthorizedSignature.png"
          );

        const acknowledgmentSignature =
          attachments.find(
            (x: any) =>
              x.FileName ===
              "AcknowledgmentEmployeeSignature.png"
          );

        const nomineeItems = await getListItemsByParentId(
          sp,
          'GratuityNominees',
          parentId
        );

        const witnessItems = await getListItemsByParentId(
          sp,
          'GratuityWitnesses',
          parentId
        );

        // SET FORM VALUES
        await formik.setValues({

          employeeIntroName:
            item.EmployeeName || "",

          spouseExclusionDate: formatSharePointDateForInput(
            item.SpouseExclusionDate as string
          ),

          nominees:
            nomineeItems?.length
              ? nomineeItems.map((n: any) => ({
                fullNameAndAddress:
                  n.FullNameAndAddress || "",

                relationship:
                  n.Relationship || "",

                age:
                  n.Age
                    ? String(n.Age)
                    : "",

                sharePercentage:
                  n.SharePercentage
                    ? String(n.SharePercentage)
                    : ""
              }))
              : [createEmptyNominee()],

          employeeStatementNameAndAddress:
            item.EmployeeStatName || "",

          sex:
            item.Sex || "",

          religion:
            item.Religion || "",

          maritalStatus:
            item.MaritalStatus || "",

          departmentBranchSection:
            item.Department || "",

          employeeId:
            item.EmployeeId
              ? String(item.EmployeeId)
              : "",

          dateOfJoining: formatSharePointDateForInput(item.DateOfJoining as string),

          permanentAddress:
            item.PermanentAddress || "",

          place:
            item.Place || "",

          date: formatSharePointDateForInput(item.FormDate0 as string),

          employeeSignature: toAttachmentPreviewUrl(
            employeeSignature?.ServerRelativeUrl
          ),

          witnesses:
            witnessItems?.length
              ? witnessItems.map(
                (w: Record<string, unknown>, index: number) => ({
                  nameAndAddress: String(w.NameAndAddress || ''),
                  signature:
                    index === 0
                      ? toAttachmentPreviewUrl(witness1Signature?.ServerRelativeUrl)
                      : toAttachmentPreviewUrl(witness2Signature?.ServerRelativeUrl)
                })
              )
              : [
                createEmptyWitness(),
                createEmptyWitness()
              ],

          referenceNo:
            item.ReferenceNo || "",

          employerCertificateDate: formatSharePointDateForInput(
            item.EmployerCertificateDate as string
          ),

          authorizedSignature: toAttachmentPreviewUrl(
            authorizedSignature?.ServerRelativeUrl
          ),

          designation:
            item.Designation || "",

          acknowledgmentDate: formatSharePointDateForInput(
            item.AcknowledgmentDate as string
          ),

          acknowledgmentEmployeeSignature: toAttachmentPreviewUrl(
            acknowledgmentSignature?.ServerRelativeUrl
          )
        });

        setSignaturePreview({
          employeeSignature: toAttachmentPreviewUrl(
            employeeSignature?.ServerRelativeUrl
          ),
          witness0: toAttachmentPreviewUrl(witness1Signature?.ServerRelativeUrl),
          witness1: toAttachmentPreviewUrl(witness2Signature?.ServerRelativeUrl),
          authorizedSignature: toAttachmentPreviewUrl(
            authorizedSignature?.ServerRelativeUrl
          ),
          acknowledgmentEmployeeSignature: toAttachmentPreviewUrl(
            acknowledgmentSignature?.ServerRelativeUrl
          )
        });

      } catch (error) {

        console.error(
          "Load Existing Data Error",
          error
        );
      }
    };

    if (employeePFData?.ID) {
      void loadExistingData();
    }

  }, [employeePFData]);

  React.useEffect(() => {
    const autoFill = async (): Promise<void> => {
      const employeeName = getCandidateValue(employeePFData, candidateFieldNames.employeeName);
      const permanentAddress = getCandidateValue(employeePFData, candidateFieldNames.address);

      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'employeeIntroName', employeeName);
      await setFieldIfEmpty(
        formik.values,
        formik.setFieldValue,
        'employeeStatementNameAndAddress',
        [employeeName, permanentAddress].filter(Boolean).join(', ')
      );
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'sex', getCandidateValue(employeePFData, candidateFieldNames.gender));
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'departmentBranchSection', getCandidateValue(employeePFData, candidateFieldNames.department));
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'employeeId', getCandidateValue(employeePFData, candidateFieldNames.employeeId));
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'dateOfJoining', getCandidateDateValue(employeePFData, candidateFieldNames.joiningDate));
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'permanentAddress', permanentAddress);
      await setFieldIfEmpty(formik.values, formik.setFieldValue, 'designation', getCandidateValue(employeePFData, candidateFieldNames.designation));
    };

    void autoFill();
  }, [employeePFData]);




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

    pdf.save("Gratuity-Nomination.pdf");
  };

  return (
    <>
      <FormSection title="Gratuity Nomination Form" />
      <div ref={formRef} className="no-break">

        <Card className="shadow border-0 rounded-4 mx-auto">
          <Card.Body className="p-4 p-md-5">
            <div className="text-center mb-4">
              <div className="fw-bold">Form &apos;F&apos;</div>
              <small>[See sub-rule(1) of rule 6]</small>

              <h3
                className="mt-3 fw-bold text-white py-2 rounded"
                style={{ backgroundColor: '#f18200' }}
              >
                GRATUITY NOMINATION FORM Testing
              </h3>
            </div>
            <div className="mb-4">
              <p className="mb-1 fw-bold">To</p>
              <p className="mb-1">NAT IT Services Pvt Ltd</p>
              <p className="mb-0">
                Plot No 21, Sruthi Sada, Serlingampally Mandal, Gachibowli, Hyderabad,
                Telangana - 500032
              </p>
            </div>

            <Form onSubmit={formik.handleSubmit}>
              <Card className="border-0 bg-light mb-4">
                <Card.Body>
                  <div className="d-flex flex-wrap align-items-start gap-2">
                    <span className="mt-2">I, Mr./Ms./Mrs</span>
                    <div>
                      <Form.Control
                        name="employeeIntroName"
                        style={{ width: '250px' }}
                        placeholder="Full Name"
                        value={formik.values.employeeIntroName}
                        onChange={formik.handleChange}
                        onBlur={formik.handleBlur}
                      />
                      {formik.touched.employeeIntroName && formik.errors.employeeIntroName && (
                        <p className="text-danger small mb-0">
                          {formik.errors.employeeIntroName}
                        </p>
                      )}
                    </div>
                    <span className="mt-2">
                      whose particulars are given below, hereby nominate the person(s)
                      mentioned below to receive gratuity payable after my death.
                    </span>
                  </div>

                  <p className="mt-3 mb-1">
                    2. I certify that the nominee(s) are members of my family.
                  </p>

                  <p className="mb-1">
                    3. I declare that I have no family under the said Act.
                  </p>

                  <div className="d-flex flex-wrap gap-2 align-items-start">
                    <span className="mt-2">
                      4. I have excluded my spouse from my family by notice dated
                    </span>
                    <div>
                      <div style={{ width: '220px' }}>
                        <DatePickerInput
                          name="spouseExclusionDate"
                          value={formik.values.spouseExclusionDate}
                          onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'spouseExclusionDate')}
                          onBlur={formik.handleBlur}
                        />
                      </div>
                      {formik.touched.spouseExclusionDate &&
                        formik.errors.spouseExclusionDate && (
                          <p className="text-danger small mb-0">
                            {formik.errors.spouseExclusionDate}
                          </p>
                        )}
                    </div>
                  </div>

                  <p className="mt-2 mb-0">
                    5. This nomination invalidates all previous nominations.
                  </p>
                </Card.Body>
              </Card>

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
                        ]?.fullNameAndAddress?.trim() ||
                        !formik.values.nominees[
                          formik.values.nominees.length - 1
                        ]?.relationship?.trim() ||
                        !formik.values.nominees[
                          formik.values.nominees.length - 1
                        ]?.sharePercentage
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
                    ]?.fullNameAndAddress?.trim() ||
                      !formik.values.nominees[
                        formik.values.nominees.length - 1
                      ]?.relationship?.trim() ||
                      !formik.values.nominees[
                        formik.values.nominees.length - 1
                      ]?.sharePercentage) && (
                        <p className="text-danger small mt-1 mb-0">
                          Please fill current nominee details before adding another nominee.
                        </p>
                      )}
                  </div>
                )}
              </div>


              <Table bordered responsive className="mt-3 text-center align-middle">
                <thead className="table-light">
                  <tr>
                    <th>S.No</th>
                    <th>Full Name & Address</th>
                    <th>Relationship</th>
                    <th>Age</th>
                    <th>Share %</th>
                  </tr>
                </thead>

                <tbody>
                  {formik.values.nominees.map((nominee, index) => {
                    const touchedNominee = formik.touched.nominees?.[index];
                    const errorNominee = formik.errors.nominees?.[index];

                    return (
                      <tr key={index}>
                        <td>{index + 1}</td>

                        <td>
                          <Form.Control
                            as="textarea"
                            rows={1}
                            name={`nominees[${index}].fullNameAndAddress`}
                            placeholder="Full Name & Address"
                            value={nominee.fullNameAndAddress}
                            onChange={(e) => {
                              formik.handleChange(e);
                              e.target.style.height = "auto";
                              e.target.style.height = e.target.scrollHeight + "px";
                            }}
                            onBlur={formik.handleBlur}
                            style={{ overflow: "hidden", resize: "none" }}
                          />
                          {typeof touchedNominee === 'object' &&
                            touchedNominee?.fullNameAndAddress &&
                            typeof errorNominee === 'object' &&
                            errorNominee?.fullNameAndAddress && (
                              <p className="text-danger small mb-0 text-start">
                                {errorNominee.fullNameAndAddress}
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
                          <Form.Control
                            name={`nominees[${index}].age`}
                            placeholder="Age"
                            value={nominee.age}
                            onChange={formik.handleChange}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedNominee === 'object' &&
                            touchedNominee?.age &&
                            typeof errorNominee === 'object' &&
                            errorNominee?.age && (
                              <p className="text-danger small mb-0 text-start">
                                {errorNominee.age}
                              </p>
                            )}
                        </td>
                        <td>
                          <div className="d-flex gap-2 align-items-start">
                            <div className="w-100">
                              <Form.Control
                                name={`nominees[${index}].sharePercentage`}
                                placeholder="%"
                                value={nominee.sharePercentage}
                                onChange={(e) => {
                                  let value = e.target.value.replace(/[^0-9]/g, '');

                                  // Prevent above 100 for single field
                                  if (Number(value) > 100) {
                                    value = '100';
                                  }

                                  // Clone nominees
                                  const updatedNominees = [...formik.values.nominees];
                                  updatedNominees[index].sharePercentage = value;

                                  // Calculate total
                                  const total = updatedNominees.reduce(
                                    (sum, item) => sum + Number(item.sharePercentage || 0),
                                    0
                                  );

                                  if (total <= 100) {
                                    formik.setFieldError(
                                      `nominees[${index}].sharePercentage`,
                                      ''
                                    );

                                    formik
                                      .setFieldValue(
                                        `nominees[${index}].sharePercentage`,
                                        value
                                      )
                                      .catch(() => undefined);
                                  } else {
                                    formik.setFieldError(
                                      `nominees[${index}].sharePercentage`,
                                      'Total share percentage cannot exceed 100%'
                                    );
                                  }
                                }}
                                onBlur={formik.handleBlur}
                              />

                              {typeof touchedNominee === 'object' &&
                                touchedNominee?.sharePercentage &&
                                typeof errorNominee === 'object' &&
                                errorNominee?.sharePercentage && (
                                  <p className="text-danger small mb-0 text-start">
                                    {errorNominee.sharePercentage}
                                  </p>
                                )}
                            </div>

                            {/* Remove Button */}
                            {formik.values.nominees.length > 1 && (
                              <Button
                                type="button"
                                variant="danger"
                                onClick={() => {
                                  const updatedNominees = formik.values.nominees.filter(
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

              <h5 style={headerStyle} className="mt-4">
                Employee Statement
              </h5>

              <Row className="g-3 mt-2">
                <Col md={12}>
                  <Form.Label>Full Name & Permanent Address</Form.Label>
                  <Form.Control
                    name="employeeStatementNameAndAddress"
                    value={formik.values.employeeStatementNameAndAddress}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.employeeStatementNameAndAddress &&
                    formik.errors.employeeStatementNameAndAddress && (
                      <p className="text-danger small mb-0">
                        {formik.errors.employeeStatementNameAndAddress}
                      </p>
                    )}
                </Col>

                <Col md={4}>
                  <Form.Label>Sex</Form.Label>
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

                <Col md={4}>
                  <Form.Label>Religion</Form.Label>
                  <Form.Select
                    name="religion"
                    value={formik.values.religion}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  >
                    <option value="">Select</option>

                    <option>Hindu </option>
                    <option>Muslim </option>
                    <option>Christian</option>
                  </Form.Select>
                  {formik.touched.religion && formik.errors.religion && (
                    <p className="text-danger small mb-0">{formik.errors.religion}</p>
                  )}
                </Col>

                <Col md={4}>
                  <Form.Label>Marital Status</Form.Label>
                  <Form.Select
                    name="maritalStatus"
                    value={formik.values.maritalStatus}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  >
                    <option value="">Select</option>
                    <option>Single</option>
                    <option>Married </option>
                    <option>divorce</option>
                    <option>widow</option>
                  </Form.Select>
                  {formik.touched.maritalStatus && formik.errors.maritalStatus && (
                    <p className="text-danger small mb-0">{formik.errors.maritalStatus}</p>
                  )}
                </Col>

                <Col md={6}>
                  <Form.Label>Department / Branch / Section</Form.Label>
                  <Form.Control
                    name="departmentBranchSection"
                    value={formik.values.departmentBranchSection}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.departmentBranchSection &&
                    formik.errors.departmentBranchSection && (
                      <p className="text-danger small mb-0">
                        {formik.errors.departmentBranchSection}
                      </p>
                    )}
                </Col>
                {isAdminEdit && (
                  <Col md={3}>
                    <Form.Label>EMP ID</Form.Label>
                    <Form.Control
                      name="employeeId"
                      value={formik.values.employeeId}
                      onChange={formik.handleChange}
                      onBlur={formik.handleBlur}

                    />
                    {formik.touched.employeeId && formik.errors.employeeId && (
                      <p className="text-danger small mb-0">{formik.errors.employeeId}</p>
                    )}
                  </Col>
                )}


                <Col md={3}>
                  <Form.Label>Date of Joining</Form.Label>
                  <DatePickerInput
                    name="dateOfJoining"
                    value={formik.values.dateOfJoining}
                    onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'dateOfJoining')}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.dateOfJoining && formik.errors.dateOfJoining && (
                    <p className="text-danger small mb-0">{formik.errors.dateOfJoining}</p>
                  )}
                </Col>

                <Col md={12}>
                  <Form.Label>Permanent Address</Form.Label>
                  <Form.Control
                    as="textarea"
                    rows={1}
                    name="permanentAddress"
                    placeholder="Permanent Address"
                    value={formik.values.permanentAddress}
                    onChange={(e) => {
                      formik.handleChange(e);
                      e.target.style.height = "auto";
                      e.target.style.height = e.target.scrollHeight + "px";
                    }}
                    onBlur={formik.handleBlur}
                    style={{ overflow: "hidden", resize: "none" }}
                  />

                  {formik.touched.permanentAddress && formik.errors.permanentAddress && (
                    <p className="text-danger small mb-0">{formik.errors.permanentAddress}</p>
                  )}
                </Col>
              </Row>

              <Row className="g-3 mt-4">
                <Col md={6}>
                  <Form.Label>Place</Form.Label>
                  <Form.Control
                    name="place"
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
                    value={
                      signaturePreview.employeeSignature ||
                      formik.values.employeeSignature
                    }
                    onChange={(file) => {

                      formik
                        .setFieldValue("employeeSignature", file)
                        .catch(() => undefined);

                      formik
                        .setFieldValue(
                          "acknowledgmentEmployeeSignature",
                          file
                        )
                        .catch(() => undefined);

                      setSignaturePreview((prev: any) => ({
                        ...prev,
                        employeeSignature: URL.createObjectURL(file),
                        acknowledgmentEmployeeSignature:
                          URL.createObjectURL(file)
                      }));
                    }}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.employeeSignature && formik.errors.employeeSignature && (
                    <p className="text-danger small mb-0">{formik.errors.employeeSignature}</p>
                  )}
                </Col>
              </Row>

              <h5 style={headerStyle} className="mt-5">
                Declaration by Witnesses
              </h5>

              <Table bordered responsive className="mt-3 text-center align-middle">
                <thead className="table-light">
                  <tr>
                    <th>Witness Name & Address</th>
                    <th>Signature</th>
                  </tr>
                </thead>

                <tbody>
                  {formik.values.witnesses.map((witness, index) => {
                    const touchedWitness = formik.touched.witnesses?.[index];
                    const errorWitness = formik.errors.witnesses?.[index];

                    return (
                      <tr key={index}>
                        <td>
                          <Form.Control
                            name={`witnesses[${index}].nameAndAddress`}
                            placeholder={`Witness ${index + 1}`}
                            value={witness.nameAndAddress}
                            onChange={formik.handleChange}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedWitness === 'object' &&
                            touchedWitness?.nameAndAddress &&
                            typeof errorWitness === 'object' &&
                            errorWitness?.nameAndAddress && (
                              <p className="text-danger small mb-0 text-start">
                                {errorWitness.nameAndAddress}
                              </p>
                            )}
                        </td>

                        <td>
                          <SignatureUpload
                            name={`witnesses[${index}].signature`}
                            value={
                              signaturePreview[`witness${index}`] ||
                              witness.signature
                            }
                            onChange={(file) => {

                              formik
                                .setFieldValue(
                                  `witnesses[${index}].signature`,
                                  file
                                )
                                .catch(() => undefined);

                              setSignaturePreview((prev: any) => ({
                                ...prev,
                                [`witness${index}`]:
                                  URL.createObjectURL(file)
                              }));
                            }}
                            onBlur={formik.handleBlur}
                          />
                          {typeof touchedWitness === 'object' &&
                            touchedWitness?.signature &&
                            typeof errorWitness === 'object' &&
                            errorWitness?.signature && (
                              <p className="text-danger small mb-0 text-start">
                                {errorWitness.signature}
                              </p>
                            )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </Table>

              <h5 style={headerStyle} className="mt-5">
                Certificate by Employer
              </h5>

              <Row className="g-3 mt-2">
                <Col md={6}>
                  <Form.Label>Reference No</Form.Label>
                  <Form.Control
                    name="referenceNo"
                    value={formik.values.referenceNo}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.referenceNo && formik.errors.referenceNo && (
                    <p className="text-danger small mb-0">{formik.errors.referenceNo}</p>
                  )}
                </Col>

                <Col md={6}>
                  <Form.Label>Date</Form.Label>
                  <DatePickerInput
                    name="employerCertificateDate"
                    value={formik.values.employerCertificateDate}
                    onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'employerCertificateDate')}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.employerCertificateDate &&
                    formik.errors.employerCertificateDate && (
                      <p className="text-danger small mb-0">
                        {formik.errors.employerCertificateDate}
                      </p>
                    )}
                </Col>

                <Col md={6}>
                  <Form.Label>Authorized Signature</Form.Label>
                  <SignatureUpload
                    name="authorizedSignature"
                    value={
                      signaturePreview.authorizedSignature ||
                      formik.values.authorizedSignature
                    }
                    onChange={(file) => {

                      formik
                        .setFieldValue(
                          "authorizedSignature",
                          file
                        )
                        .catch(() => undefined);

                      setSignaturePreview((prev: any) => ({
                        ...prev,
                        authorizedSignature:
                          URL.createObjectURL(file)
                      }));
                    }}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.authorizedSignature && formik.errors.authorizedSignature && (
                    <p className="text-danger small mb-0">{formik.errors.authorizedSignature}</p>
                  )}
                </Col>

                <Col md={6}>
                  <Form.Label>Designation</Form.Label>
                  <Form.Control
                    name="designation"
                    value={formik.values.designation}
                    onChange={formik.handleChange}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.designation && formik.errors.designation && (
                    <p className="text-danger small mb-0">{formik.errors.designation}</p>
                  )}
                </Col>
              </Row>

              <h5 style={headerStyle} className="mt-5">
                Acknowledgment by Employee
              </h5>

              <Row className="g-3 mt-2">
                <Col md={6}>
                  <Form.Label>Date</Form.Label>
                  <DatePickerInput
                    name="acknowledgmentDate"
                    value={formik.values.acknowledgmentDate}
                    onValueChange={createDateValueChangeHandler(formik.setFieldValue, 'acknowledgmentDate')}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.acknowledgmentDate && formik.errors.acknowledgmentDate && (
                    <p className="text-danger small mb-0">{formik.errors.acknowledgmentDate}</p>
                  )}
                </Col>

                <Col md={6}>
                  <Form.Label>Employee Signature</Form.Label>
                  {/* <SignatureUpload
                    name="acknowledgmentEmployeeSignature"
                    value={formik.values.acknowledgmentEmployeeSignature}
                    onChange={(value) => {
                      formik.setFieldValue('acknowledgmentEmployeeSignature', value).catch(() => undefined);
                      formik.setFieldValue('employeeSignature', value).catch(() => undefined);
                      onEmployeeSignatureChange?.(value);
                    }}
                    onBlur={formik.handleBlur}
                  /> */}
                  <SignatureUpload
                    name="acknowledgmentEmployeeSignature"
                    value={
                      signaturePreview.acknowledgmentEmployeeSignature ||
                      formik.values.acknowledgmentEmployeeSignature
                    }
                    onChange={(file) => {

                      formik
                        .setFieldValue(
                          "acknowledgmentEmployeeSignature",
                          file
                        )
                        .catch(() => undefined);

                      formik
                        .setFieldValue(
                          "employeeSignature",
                          file
                        )
                        .catch(() => undefined);

                      setSignaturePreview((prev: any) => ({
                        ...prev,
                        acknowledgmentEmployeeSignature:
                          URL.createObjectURL(file),
                        employeeSignature:
                          URL.createObjectURL(file)
                      }));
                    }}
                    onBlur={formik.handleBlur}
                  />
                  {formik.touched.acknowledgmentEmployeeSignature &&
                    formik.errors.acknowledgmentEmployeeSignature && (
                      <p className="text-danger small mb-0">
                        {formik.errors.acknowledgmentEmployeeSignature}
                      </p>
                    )}
                </Col>
              </Row>

              {readOnly && (
                <div className="alert alert-success mt-4 mb-0">
                  Your forms have been finally submitted and can no longer be edited.
                </div>
              )}
              <div className="text-center mt-5">
                {!readOnly && (
                  <Button
                    type="submit"
                    size="lg"
                    style={{
                      backgroundColor: '#f18200',
                      border: 'none',
                      minWidth: '200px'
                    }}
                  >
                    {submitButtonLabel}
                  </Button>
                )}
                {isAdminEdit && (
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

export default GratuityNominationForm;
