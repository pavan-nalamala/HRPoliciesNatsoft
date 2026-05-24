import * as React from "react";
import {
    Container,
    Card,
    Form,
    Row,
    Col,
    Button,
    Alert,
} from "react-bootstrap";
import { useFormik } from "formik";
import * as Yup from "yup";
import moment from "moment";
import type { ISequentialFormProps } from "./ISequentialFormProps";
import {
    createDateValueChangeHandler,
    createMatchingDateValidation,
    createDateValidation,
    DatePickerInput,
    formatDateForDisplay
} from "./dateFieldUtils";
import { getSP } from "../../../../pnpjsConfig";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import {
    candidateFieldNames,
    getCandidateDateValue,
    getCandidateValue,
    setFieldIfEmpty
} from "./candidateAutoFillUtils";

type PFDeclarationValues = {
    employeeName: string;
    dateOfBirth: string;
    fatherOrSpouseName: string;
    relationType: string;
    gender: string;
    maritalStatus: string;
    email: string;
    mobileNo: string;
    epf1952: string;
    eps1995: string;
    uan: string;
    previousPf: string;
    exitPreviousEmployment: string;
    schemeCertificateNo: string;
    ppo: string;
    internationalWorker: string;
    countryOrigin: string;
    passportNo: string;
    passportValidity: string;
    educationalQualification: string;
    speciallyAbled: string;
    disabilityCategory: string;
    bankAccNo: string;
    ifscCode: string;
    aadharNo: string;
    doHavePan: string;
    pan: string;
    place: string;
    declarationAccepted: boolean;
    employerMemberName: string;
    employerJoinDate: string;
    employerPfMemberId: string;
    employerUan: string;
    employerKycPending: boolean;
    employerKycUploadedNotApproved: boolean;
    employerKycApproved: boolean;
    employerTransferApproved: boolean;
    employerPhysicalClaim: boolean;
    employerDate: string;
};

const initialValues: PFDeclarationValues = {
    employeeName: "",
    dateOfBirth: "",
    fatherOrSpouseName: "",
    relationType: "",
    gender: "",
    maritalStatus: "",
    email: "",
    mobileNo: "",
    epf1952: "",
    eps1995: "",
    uan: "",
    previousPf: "",
    exitPreviousEmployment: "",
    schemeCertificateNo: "",
    ppo: "",
    internationalWorker: "",
    countryOrigin: "",
    passportNo: "",
    passportValidity: "",
    educationalQualification: "",
    speciallyAbled: "",
    disabilityCategory: "",
    bankAccNo: "",
    ifscCode: "",
    aadharNo: "",
    doHavePan: "",
    pan: "",
    place: "",
    declarationAccepted: false,
    employerMemberName: "",
    employerJoinDate: "",
    employerPfMemberId: "",
    employerUan: "",
    employerKycPending: false,
    employerKycUploadedNotApproved: false,
    employerKycApproved: false,
    employerTransferApproved: false,
    employerPhysicalClaim: false,
    employerDate: "",
};


export default function PFDeclaration({ onComplete, spHttpClient, siteUrl, context, isSubmitted, employeePFData, sharedDateOfBirth }: ISequentialFormProps): JSX.Element {
    const formRef = React.useRef<HTMLDivElement>(null);

    const commonDateOfBirth =
        sharedDateOfBirth || getCandidateDateValue(employeePFData, candidateFieldNames.dateOfBirth);
    const commonDateOfBirthMessage = commonDateOfBirth
        ? `Date of Birth must match the employee Date of Birth (${commonDateOfBirth})`
        : "Date of Birth must match the employee Date of Birth";

    const validationSchema = Yup.object().shape({
        employeeName: Yup.string().required("Employee Name is required"),
        dateOfBirth: createMatchingDateValidation(
            createDateValidation("Date of Birth is required", "Date of Birth must be in DD/MM/YYYY format"),
            commonDateOfBirth,
            commonDateOfBirthMessage
        ),
        relationType: Yup.string().required("Please select Father or Spouse"),
        fatherOrSpouseName: Yup.string().required("Father / Spouse is required"),
        gender: Yup.string().required("Gender is required"),
        maritalStatus: Yup.string().required("Marital Status is required"),
        epf1952: Yup.string().required("Earlier Member EPF 1952 is required"),
        eps1995: Yup.string().required("Earlier Member EPS 1995 is required"),
        internationalWorker: Yup.string().required("International Worker is required"),
        educationalQualification: Yup.string().required("Educational Qualification is required"),
        speciallyAbled: Yup.string().required("Specially Abled is required"),
        disabilityCategory: Yup.string().when("speciallyAbled", {
            is: "Yes",
            then: (schema) => schema.required("Disability Category is required"),
            otherwise: (schema) => schema.notRequired(),
        }),
        countryOrigin: Yup.string().when("internationalWorker", {
            is: "Yes",
            then: (schema) => schema.required("Country Origin is required"),
            otherwise: (schema) => schema.notRequired(),
        }),
        passportNo: Yup.string().when("internationalWorker", {
            is: "Yes",
            then: (schema) => schema.required("Passport No is required"),
            otherwise: (schema) => schema.notRequired(),
        }),
        passportValidity: Yup.string().when("internationalWorker", {
            is: "Yes",
            then: () => createDateValidation("Passport Validity is required", "Passport Validity must be in DD/MM/YYYY format"),
            otherwise: (schema) => schema.notRequired(),
        }),
        pan: Yup.string().when("doHavePan", {
            is: "Yes",
            then: (schema) =>
                schema
                    .required("PAN is required")
                    .matches(/^[A-Z]{5}[0-9]{4}[A-Z]$/, "Invalid PAN"),
            otherwise: (schema) => schema.notRequired(),
        }),
        declarationAccepted: Yup.boolean().oneOf(
            [true],
            "Please acknowledge the information provided is correct"
        ),
    });

    const formik = useFormik<PFDeclarationValues>({
        initialValues,
        validationSchema,
        // onSubmit: async (values) => {

        //     try {

        //         const sp = getSP(context);

        //         await sp.web.lists
        //             .getByTitle("EPFDeclarationForm11")
        //             .items.add({

        //                 Title: values.employeeName,

        //                 EmployeeName: values.employeeName,

        //                 DateOfBirth:
        //                     values.dateOfBirth
        //                         ? moment(values.dateOfBirth, "DD/MM/YYYY").toISOString()
        //                         : null,

        //                 FatherOrSpouse: values.relationType,

        //                 FatherSpouseName: values.fatherOrSpouseName,

        //                 Gender: values.gender,

        //                 MaritalStatus: values.maritalStatus,

        //                 Email: values.email,

        //                 Mobile: values.mobileNo,

        //                 // SharePoint YES/NO columns
        //                 EarlierMemberEPF1952:
        //                     values.epf1952 === "Yes",

        //                 EarlierMemberEPS1995:
        //                     values.eps1995 === "Yes",

        //                 InternationalWorker:
        //                     values.internationalWorker === "Yes",

        //                 CountryOrigin:
        //                     values.countryOrigin || "",

        //                 PassportNo:
        //                     values.passportNo || "",

        //                 PassportValidity:
        //                     values.passportValidity
        //                         ? moment(values.passportValidity, "DD/MM/YYYY").toISOString()
        //                         : null,

        //                 EducationalQualification:
        //                     values.educationalQualification,

        //                 SpeciallyAbled:
        //                     values.speciallyAbled === "Yes",

        //                 DisabilityCategory:
        //                     values.disabilityCategory || "",

        //                 BankAccountNo:
        //                     values.bankAccNo,

        //                 IFSCCode:
        //                     values.ifscCode,

        //                 AadhaarNo:
        //                     values.aadharNo,

        //                 DoHavePan:
        //                     values.doHavePan === "Yes",

        //                 PANNo:
        //                     values.pan || "",

        //                 Place:
        //                     values.place,

        //                 DeclarationAccepted:
        //                     values.declarationAccepted || false,

        //                 EmployeeDeclarationDate:
        //                     new Date().toISOString(),

        //                 PresentEmployerName:
        //                     values.employerMemberName || "",

        //                 DateOfJoining:
        //                     values.employerJoinDate
        //                         ? moment(values.employerJoinDate, "DD/MM/YYYY").toISOString()
        //                         : null,

        //                 PFMemberID:
        //                     values.employerPfMemberId || "",

        //                 UAN:
        //                     values.employerUan || values.uan || "",

        //                 PreviousPFDetails:
        //                     values.previousPf || "",

        //                 EmployerKycPending:
        //                     values.employerKycPending || false,

        //                 UANUploadedNotApproved:
        //                     values.employerKycUploadedNotApproved || false,

        //                 UANApprovedWithDSC:
        //                     values.employerKycApproved || false,

        //                 PreviousPFTransferred:
        //                     values.employerTransferApproved || false,

        //                 PhysicalClaimRequired:
        //                     values.employerPhysicalClaim || false,

        //                 EmployerDeclarationDate:
        //                     values.employerDate
        //                         ? moment(values.employerDate, "DD/MM/YYYY").toISOString()
        //                         : null,

        //                 EmployerName:
        //                     values.employerMemberName || "",

        //                 SubmittedBy:
        //                     context?.pageContext.user.displayName || "",

        //                 SubmittedEmail:
        //                     context?.pageContext.user.email || "",

        //                 SubmittedTime:
        //                     new Date().toISOString()
        //             });

        //         onComplete?.();

        //     } catch (error: any) {

        //         console.error("Submit Error => ", error);


        //         alert(error?.message || "Submission failed");
        //     }
        // }

        onSubmit: async (values) => {

            try {

                const sp = getSP(context);

                // GENERATE UNIQUE TOKEN
                // const token = crypto.randomUUID();
                const token =
                    Date.now().toString() +
                    Math.random()
                        .toString(36)
                        .substring(2, 10);

                const deepLink =
                    `${window.location.origin}` +
                    `${context?.pageContext.web.serverRelativeUrl}` +
                    `/SitePages/PFForms.aspx#/admin-edit/${token}?mode=hr`;
                // SAVE FORM
                const response = await sp.web.lists
                    .getByTitle("EPFDeclarationForm11")
                    .items.add({

                        Title: values.employeeName,

                        EmployeeName: values.employeeName,

                        DateOfBirth:
                            values.dateOfBirth
                                ? moment(
                                    values.dateOfBirth,
                                    "DD/MM/YYYY"
                                ).toISOString()
                                : null,

                        FatherOrSpouse:
                            values.relationType,

                        FatherSpouseName:
                            values.fatherOrSpouseName,

                        Gender:
                            values.gender,

                        MaritalStatus:
                            values.maritalStatus,

                        Email:
                            values.email,

                        Mobile:
                            values.mobileNo,

                        EarlierMemberEPF1952:
                            values.epf1952 === "Yes",

                        EarlierMemberEPS1995:
                            values.eps1995 === "Yes",

                        InternationalWorker:
                            values.internationalWorker === "Yes",

                        CountryOrigin:
                            values.countryOrigin || "",

                        PassportNo:
                            values.passportNo || "",

                        PassportValidity:
                            values.passportValidity
                                ? moment(
                                    values.passportValidity,
                                    "DD/MM/YYYY"
                                ).toISOString()
                                : null,

                        EducationalQualification:
                            values.educationalQualification,

                        SpeciallyAbled:
                            values.speciallyAbled === "Yes",

                        DisabilityCategory:
                            values.disabilityCategory || "",

                        BankAccountNo:
                            values.bankAccNo,

                        IFSCCode:
                            values.ifscCode,

                        AadhaarNo:
                            values.aadharNo,

                        DoHavePan:
                            values.doHavePan === "Yes",

                        PANNo:
                            values.pan || "",

                        Place:
                            values.place,

                        DeclarationAccepted:
                            values.declarationAccepted || false,

                        EmployeeDeclarationDate:
                            new Date().toISOString(),

                        PresentEmployerName:
                            values.employerMemberName || "",

                        DateOfJoining:
                            values.employerJoinDate
                                ? moment(
                                    values.employerJoinDate,
                                    "DD/MM/YYYY"
                                ).toISOString()
                                : null,

                        PFMemberID:
                            values.employerPfMemberId || "",

                        UAN:
                            values.employerUan ||
                            values.uan ||
                            "",

                        PreviousPFDetails:
                            values.previousPf || "",

                        EmployerKycPending:
                            values.employerKycPending || false,

                        UANUploadedNotApproved:
                            values.employerKycUploadedNotApproved || false,

                        UANApprovedWithDSC:
                            values.employerKycApproved || false,

                        PreviousPFTransferred:
                            values.employerTransferApproved || false,

                        PhysicalClaimRequired:
                            values.employerPhysicalClaim || false,

                        EmployerDeclarationDate:
                            values.employerDate
                                ? moment(
                                    values.employerDate,
                                    "DD/MM/YYYY"
                                ).toISOString()
                                : null,

                        EmployerName:
                            values.employerMemberName || "",

                        SubmittedBy:
                            context?.pageContext.user.displayName || "",

                        SubmittedEmail:
                            context?.pageContext.user.email || "",

                        SubmittedTime:
                            new Date().toISOString(),

                        // HR FLOW
                        HRStatus: "Pending",

                        AdminDeepLink:
                            deepLink,

                        // EMPLOYEE UNIQUE ID
                        can_id:
                            String(employeePFData?.ID),

                        // UNIQUE TOKEN
                        UniqueToken:
                            token
                    });

                console.log(
                    "HR Link:",
                    deepLink
                );
                console.log("response", response);


                alert(`
Form Submitted Successfully

HR Link:
${deepLink}
`);

                // OPTIONAL
                window.open(
                    deepLink,
                    "_blank"
                );

                onComplete?.();

            } catch (error: any) {

                console.error(
                    "Submit Error => ",
                    error
                );

                alert(
                    error?.message ||
                    "Submission failed"
                );
            }
        }


    });

    React.useEffect(() => {
        const autoFill = async (): Promise<void> => {
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "employeeName", getCandidateValue(employeePFData, candidateFieldNames.employeeName));
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "dateOfBirth", getCandidateDateValue(employeePFData, candidateFieldNames.dateOfBirth));
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "fatherOrSpouseName", getCandidateValue(employeePFData, candidateFieldNames.fatherOrHusbandName));
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "gender", getCandidateValue(employeePFData, candidateFieldNames.gender));
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "maritalStatus", getCandidateValue(employeePFData, candidateFieldNames.maritalStatus));
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "email", getCandidateValue(employeePFData, candidateFieldNames.email));
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "mobileNo", getCandidateValue(employeePFData, candidateFieldNames.mobile));
            await setFieldIfEmpty(formik.values, formik.setFieldValue, "pan", getCandidateValue(employeePFData, candidateFieldNames.panNo));
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

        pdf.save("PF-Declaration.pdf");
    };

    const handleExclusiveCheckboxChange = (
        selectedField:
            | "employerKycPending"
            | "employerKycUploadedNotApproved"
            | "employerKycApproved"
            | "employerTransferApproved"
            | "employerPhysicalClaim",
        groupedFields: Array<
            | "employerKycPending"
            | "employerKycUploadedNotApproved"
            | "employerKycApproved"
            | "employerTransferApproved"
            | "employerPhysicalClaim"
        >
    ) => (event: React.ChangeEvent<HTMLInputElement>) => {
        const isChecked = event.target.checked;

        groupedFields.forEach((fieldName) => {
            formik.setFieldValue(fieldName, fieldName === selectedField ? isChecked : false).catch(() => undefined);
        });
    };

    const currentDate = formatDateForDisplay(new Date());

    return (
        <Container fluid className="py-4 bg-light no-break" ref={formRef} >
            <Card
                className="mx-auto shadow border-0 rounded-4"

            >
                <Card.Header
                    className="text-white text-center fw-bold fs-4"
                    style={{ background: "#f18200" }}
                >
                    Form No.11 Declaration Form
                </Card.Header>
                <Card.Body className="p-4">

                    <h5 className="text-white p-2 rounded" style={{ backgroundColor: "rgb(241, 130, 0)" }}>Employees Provident Fund Organisation Declaration Form</h5>

                    <Form onSubmit={formik.handleSubmit}>
                        <Row className="g-3">
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Employee Name *</Form.Label>
                                    <Form.Control
                                        name="employeeName"
                                        value={formik.values.employeeName}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    />
                                    {formik.touched.employeeName && formik.errors.employeeName && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.employeeName}
                                        </p>
                                    )}
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Date of Birth *</Form.Label>
                                    <DatePickerInput
                                        name="dateOfBirth"
                                        value={formik.values.dateOfBirth}
                                        onValueChange={createDateValueChangeHandler(formik.setFieldValue, "dateOfBirth")}
                                        onBlur={formik.handleBlur}
                                    />
                                    {formik.touched.dateOfBirth && formik.errors.dateOfBirth && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.dateOfBirth}
                                        </p>
                                    )}
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Label>Father / Spouse *</Form.Label>
                                <div className="d-flex gap-3 mb-2">
                                    <Form.Check
                                        type="radio"
                                        name="relationType"
                                        label="Father"
                                        value="Father"
                                        checked={formik.values.relationType === "Father"}
                                        onChange={formik.handleChange}
                                        className="customCheckbox "
                                    />
                                    <Form.Check
                                        type="radio"
                                        name="relationType"
                                        label="Spouse"
                                        value="Spouse"
                                        className="customCheckbox "
                                        checked={formik.values.relationType === "Spouse"}
                                        onChange={formik.handleChange}
                                    />
                                </div>
                                {formik.values.relationType && (
                                    <Form.Control
                                        name="fatherOrSpouseName"
                                        value={formik.values.fatherOrSpouseName}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                        placeholder={`Enter ${formik.values.relationType} Name`}
                                    />
                                )}
                                {((formik.touched.relationType && formik.errors.relationType) ||
                                    (formik.touched.fatherOrSpouseName &&
                                        formik.errors.fatherOrSpouseName)) && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.relationType || formik.errors.fatherOrSpouseName}
                                        </p>
                                    )}
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Gender *</Form.Label>
                                    <Form.Select
                                        name="gender"
                                        value={formik.values.gender}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    >
                                        <option value="">Select</option>
                                        <option>Male</option>
                                        <option>Female</option>
                                        <option>Others</option>
                                    </Form.Select>
                                    {formik.touched.gender && formik.errors.gender && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.gender}
                                        </p>
                                    )}
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Marital Status *</Form.Label>
                                    <Form.Select
                                        name="maritalStatus"
                                        value={formik.values.maritalStatus}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    >
                                        <option value="">Select</option>

                                        <option>Single</option>
                                        <option>Married</option>
                                        <option>divorce</option>
                                        <option>widow</option>
                                    </Form.Select>
                                    {formik.touched.maritalStatus && formik.errors.maritalStatus && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.maritalStatus}
                                        </p>
                                    )}
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Email</Form.Label>
                                    <Form.Control
                                        name="email"
                                        value={formik.values.email}
                                        onChange={formik.handleChange}
                                    />
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Mobile</Form.Label>
                                    <Form.Control
                                        name="mobileNo"
                                        value={formik.values.mobileNo}
                                        onChange={(e) =>
                                            formik.setFieldValue(
                                                "mobileNo",
                                                e.target.value.replace(/\D/g, "")
                                            ).catch(() => undefined)
                                        }
                                    />
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Earlier Member EPF 1952 *</Form.Label>
                                    <Form.Select
                                        name="epf1952"
                                        value={formik.values.epf1952}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    >
                                        <option value="">Select</option>
                                        <option>Yes</option>
                                        <option>No</option>
                                    </Form.Select>
                                    {formik.touched.epf1952 && formik.errors.epf1952 && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.epf1952}
                                        </p>
                                    )}
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Earlier Member EPS 1995 *</Form.Label>
                                    <Form.Select
                                        name="eps1995"
                                        value={formik.values.eps1995}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    >
                                        <option value="">Select</option>
                                        <option>Yes</option>
                                        <option>No</option>
                                    </Form.Select>
                                    {formik.touched.eps1995 && formik.errors.eps1995 && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.eps1995}
                                        </p>
                                    )}
                                </Form.Group>
                            </Col>
                        </Row>

                        {(formik.values.epf1952 === "Yes" || formik.values.eps1995 === "Yes") && (
                            <>
                                <h5 className="mt-4">Previous Employment Details</h5>
                                <Row className="g-3">
                                    <Col md={4}>
                                        <Form.Control
                                            name="uan"
                                            placeholder="UAN"
                                            value={formik.values.uan}
                                            onChange={(e) =>
                                                formik.setFieldValue(
                                                    "uan",
                                                    e.target.value.replace(/\D/g, "").slice(0, 12)
                                                ).catch(() => undefined)
                                            }
                                        />
                                    </Col>
                                    <Col md={4}>
                                        <Form.Control
                                            name="previousPf"
                                            placeholder="Previous PF"
                                            value={formik.values.previousPf}
                                            onChange={formik.handleChange}
                                        />
                                    </Col>
                                    <Col md={4}>
                                        <DatePickerInput
                                            name="exitPreviousEmployment"
                                            value={formik.values.exitPreviousEmployment}
                                            onValueChange={createDateValueChangeHandler(formik.setFieldValue, "exitPreviousEmployment")}
                                        />
                                    </Col>
                                    <Col md={6}>
                                        <Form.Control
                                            name="schemeCertificateNo"
                                            placeholder="Scheme Certificate No"
                                            value={formik.values.schemeCertificateNo}
                                            onChange={formik.handleChange}
                                        />
                                    </Col>
                                    <Col md={6}>
                                        <Form.Control
                                            name="ppo"
                                            placeholder="PPO No"
                                            value={formik.values.ppo}
                                            onChange={formik.handleChange}
                                        />
                                    </Col>
                                </Row>
                            </>
                        )}

                        <Row className="g-3 mt-3">
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>International Worker *</Form.Label>
                                    <Form.Select
                                        name="internationalWorker"
                                        value={formik.values.internationalWorker}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    >
                                        <option value="">Select</option>
                                        <option>Yes</option>
                                        <option>No</option>
                                    </Form.Select>
                                    {formik.touched.internationalWorker &&
                                        formik.errors.internationalWorker && (
                                            <p className="text-danger small mb-0">
                                                {formik.errors.internationalWorker}
                                            </p>
                                        )}
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Educational Qualification *</Form.Label>
                                    <Form.Select
                                        name="educationalQualification"
                                        value={formik.values.educationalQualification}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    >
                                        <option value="">Select</option>
                                        {/* <option>Illiterate</option>
                                        <option>Non-Matric</option>
                                        <option>Matric</option> */}
                                        <option>Senior Secondary</option>
                                        <option>Bachelor Degree</option>
                                        <option>Post Graduate</option>
                                        {/* <option>Doctor</option>
                                        <option>Technical/Professional</option> */}
                                    </Form.Select>
                                    {formik.touched.educationalQualification &&
                                        formik.errors.educationalQualification && (
                                            <p className="text-danger small mb-0">
                                                {formik.errors.educationalQualification}
                                            </p>
                                        )}
                                </Form.Group>
                            </Col>
                        </Row>

                        {formik.values.internationalWorker === "Yes" && (
                            <Row className="g-3 mt-1">
                                <Col md={4}>
                                    <Form.Control
                                        name="countryOrigin"
                                        placeholder="Country Origin"
                                        value={formik.values.countryOrigin}
                                        onChange={formik.handleChange}
                                        onBlur={formik.handleBlur}
                                    />
                                    {formik.touched.countryOrigin && formik.errors.countryOrigin && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.countryOrigin}
                                        </p>
                                    )}
                                </Col>
                                <Col md={4}>
                                    <Form.Control
                                        name="passportNo"
                                        placeholder="Passport No"
                                        value={formik.values.passportNo}
                                        onChange={(e) =>
                                            formik.setFieldValue(
                                                "passportNo",
                                                e.target.value.toUpperCase().slice(0, 8)
                                            ).catch(() => undefined)
                                        }
                                        onBlur={formik.handleBlur}
                                    />
                                    {formik.touched.passportNo && formik.errors.passportNo && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.passportNo}
                                        </p>
                                    )}
                                </Col>
                                <Col md={4}>
                                    <DatePickerInput
                                        name="passportValidity"
                                        value={formik.values.passportValidity}
                                        onValueChange={createDateValueChangeHandler(formik.setFieldValue, "passportValidity")}
                                        onBlur={formik.handleBlur}
                                    />
                                    {formik.touched.passportValidity &&
                                        formik.errors.passportValidity && (
                                            <p className="text-danger small mb-0">
                                                {formik.errors.passportValidity}
                                            </p>
                                        )}
                                </Col>
                            </Row>
                        )}

                        <Row className="g-3 mt-3">
                            <Col md={6}>
                                <Form.Group>
                                    <Form.Label>Specially Abled *</Form.Label>
                                    <Form.Select
                                        name="speciallyAbled"
                                        value={formik.values.speciallyAbled}
                                        onChange={(e) => {
                                            formik.handleChange(e);
                                            formik
                                                .setFieldValue("disabilityCategory", "")
                                                .catch(() => undefined);
                                        }}
                                        onBlur={formik.handleBlur}
                                    >
                                        <option value="">Select</option>
                                        <option>Yes</option>
                                        <option>No</option>
                                    </Form.Select>
                                    {formik.touched.speciallyAbled && formik.errors.speciallyAbled && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.speciallyAbled}
                                        </p>
                                    )}
                                </Form.Group>
                            </Col>
                            <Col md={6}>
                                {formik.values.speciallyAbled === "Yes" && (
                                    <Form.Group>
                                        <Form.Label>Disability Category</Form.Label>
                                        <Form.Select
                                            name="disabilityCategory"
                                            value={formik.values.disabilityCategory}
                                            onChange={formik.handleChange}
                                            onBlur={formik.handleBlur}
                                        >
                                            <option value="">Select</option>
                                            <option>Locomotive</option>
                                            <option>Visual</option>
                                            <option>Hearing</option>
                                        </Form.Select>
                                        {formik.touched.disabilityCategory &&
                                            formik.errors.disabilityCategory && (
                                                <p className="text-danger small mb-0">
                                                    {formik.errors.disabilityCategory}
                                                </p>
                                            )}
                                    </Form.Group>
                                )}
                            </Col>
                        </Row>

                        <h5 className="text-white p-2 rounded mt-3" style={{ backgroundColor: "rgb(241, 130, 0)" }}>KYC Details</h5>
                        <Row className="g-3">
                            <Col md={4}>
                                <Form.Control
                                    name="bankAccNo"
                                    placeholder="Bank Account No"
                                    value={formik.values.bankAccNo}
                                    onChange={(e) =>
                                        formik.setFieldValue(
                                            "bankAccNo",
                                            e.target.value.replace(/\D/g, "")
                                        ).catch(() => undefined)
                                    }
                                />
                            </Col>
                            <Col md={4}>
                                <Form.Control
                                    name="ifscCode"
                                    placeholder="IFSC Code"
                                    value={formik.values.ifscCode}
                                    onChange={(e) =>
                                        formik.setFieldValue(
                                            "ifscCode",
                                            e.target.value.toUpperCase()
                                        ).catch(() => undefined)
                                    }
                                />
                            </Col>
                            <Col md={4}>
                                <Form.Control
                                    name="aadharNo"
                                    placeholder="Aadhar No"
                                    value={formik.values.aadharNo}
                                    onChange={(e) =>
                                        formik.setFieldValue(
                                            "aadharNo",
                                            e.target.value.replace(/\D/g, "").slice(0, 12)
                                        ).catch(() => undefined)
                                    }
                                    maxLength={12}
                                />
                            </Col>
                        </Row>
                        <Row className="g-3 mt-2">
                            <Col md={6}>
                                <Form.Select
                                    name="doHavePan"
                                    value={formik.values.doHavePan}
                                    onChange={formik.handleChange}
                                >
                                    <option value="">Do you have PAN?</option>
                                    <option>Yes</option>
                                    <option>No</option>
                                </Form.Select>
                            </Col>
                            <Col md={6}>
                                {formik.values.doHavePan === "Yes" && (
                                    <>
                                        <Form.Control
                                            name="pan"
                                            placeholder="PAN No"
                                            value={formik.values.pan}
                                            onChange={(e) =>
                                                formik.setFieldValue(
                                                    "pan",
                                                    e.target.value.toUpperCase().slice(0, 10)
                                                ).catch(() => undefined)
                                            }
                                            onBlur={formik.handleBlur}
                                            maxLength={10}
                                        />
                                        {formik.touched.pan && formik.errors.pan && (
                                            <p className="text-danger small mb-0">
                                                {formik.errors.pan}
                                            </p>
                                        )}
                                    </>
                                )}
                            </Col>
                        </Row>
                        <Card className="mt-4 bg-light border-0 rounded-4">
                            <Card.Body>
                                <h6 className="fw-bold">Undertaking</h6>
                                <ol className="mb-3">
                                    <li>Certified that particulars are true.</li>
                                    <li>I authorize EPFO to use Aadhaar for verification.</li>
                                    <li>Transfer previous PF details if applicable.</li>
                                    <li>I will inform employer for any changes.</li>
                                </ol>
                                <Row className="g-3">
                                    <Col md={6}>
                                        <Form.Control value={currentDate} disabled />
                                    </Col>
                                    <Col md={6}>
                                        <Form.Control
                                            name="place"
                                            placeholder="Place"
                                            value={formik.values.place}
                                            onChange={formik.handleChange}
                                        />
                                    </Col>
                                </Row>
                                <Form.Check
                                    className="mt-3 customCheckbox"

                                    name="declarationAccepted"
                                    label={
                                        <span
                                            style={{
                                                marginLeft: "6px",
                                            }}
                                        >
                                            I acknowledge the information provided is correct.
                                        </span>}
                                    checked={formik.values.declarationAccepted}
                                    onChange={formik.handleChange}
                                    onBlur={formik.handleBlur}
                                />
                                {formik.touched.declarationAccepted &&
                                    formik.errors.declarationAccepted && (
                                        <p className="text-danger small mb-0">
                                            {formik.errors.declarationAccepted}
                                        </p>
                                    )}
                                {employeePFData?.EmailID === 'hr@natit.in' && (
                                    <Card className="mt-4 ">
                                        <Card.Header
                                            className="fw-bold text-center text-uppercase text-white"
                                            style={{
                                                backgroundColor: "rgb(241, 130, 0)",
                                                fontSize: "18px",
                                                color: "#fff"
                                            }}
                                        >
                                            Declaration By Present Employer
                                        </Card.Header>

                                        <Card.Body
                                            style={{
                                                fontSize: "14px",
                                                lineHeight: "1.6"
                                            }}
                                        >
                                            <div className="mb-3">
                                                <strong>A.</strong>{" "}
                                                The member Mr./Ms./Mrs.
                                                <Form.Control
                                                    name="employerMemberName"
                                                    type="text"
                                                    className="d-inline-block mx-2"
                                                    style={{ width: "220px" }}
                                                    value={formik.values.employerMemberName}
                                                    onChange={formik.handleChange}
                                                />
                                                has joined on
                                                <span className="d-inline-block mx-2" style={{ width: "220px" }}>
                                                    <DatePickerInput
                                                        name="employerJoinDate"
                                                        value={formik.values.employerJoinDate}
                                                        onValueChange={createDateValueChangeHandler(formik.setFieldValue, "employerJoinDate")}
                                                    />
                                                </span>
                                                and has been allotted PF Member ID
                                                <Form.Control
                                                    name="employerPfMemberId"
                                                    type="text"
                                                    className="d-inline-block mx-2 mt-2"
                                                    style={{ width: "220px" }}
                                                    value={formik.values.employerPfMemberId}
                                                    onChange={formik.handleChange}
                                                />
                                            </div>

                                            <div className="mb-3">
                                                <strong>B.</strong>{" "}
                                                In case the person was earlier not a member of EPF Scheme,
                                                1952 and EPS, 1995:
                                                <ul className="mt-2 mb-2">
                                                    <li>
                                                        <strong>(Post Allotment of UAN)</strong> The UAN
                                                        allotted for the member is
                                                        <Form.Control
                                                            name="employerUan"
                                                            type="text"
                                                            className="d-inline-block mx-2"
                                                            style={{ width: "220px" }}
                                                            value={formik.values.employerUan}
                                                            onChange={formik.handleChange}
                                                        />
                                                    </li>
                                                    <li className="mt-2">
                                                        <strong>Please tick the appropriate option:</strong>
                                                    </li>
                                                </ul>

                                                <div className="ms-4">
                                                    <Form.Check
                                                        name="employerKycPending"
                                                        type="checkbox"
                                                        label={
                                                            <span
                                                                style={{
                                                                    marginLeft: "6px",
                                                                }}
                                                            >
                                                                The KYC details of the above member in the UAN database have not been uploaded
                                                            </span>}
                                                        className="mb-2 d-block customCheckbox"
                                                        checked={formik.values.employerKycPending}
                                                        onChange={handleExclusiveCheckboxChange("employerKycPending", [
                                                            "employerKycPending",
                                                            "employerKycUploadedNotApproved",
                                                            "employerKycApproved",
                                                        ])}
                                                    />

                                                    <Form.Check
                                                        name="employerKycUploadedNotApproved"
                                                        type="checkbox"
                                                        label={
                                                            <span
                                                                style={{
                                                                    marginLeft: "6px",
                                                                }}
                                                            >
                                                                Have been uploaded but not approved
                                                            </span>}

                                                        className="mb-2 customCheckbox"
                                                        checked={formik.values.employerKycUploadedNotApproved}
                                                        onChange={handleExclusiveCheckboxChange("employerKycUploadedNotApproved", [
                                                            "employerKycPending",
                                                            "employerKycUploadedNotApproved",
                                                            "employerKycApproved",
                                                        ])}
                                                    />

                                                    <Form.Check
                                                        name="employerKycApproved"
                                                        type="checkbox"
                                                        label={
                                                            <span
                                                                style={{
                                                                    marginLeft: "6px",
                                                                }}
                                                            >
                                                                Have been uploaded and approved with DSC
                                                            </span>}
                                                        className="mb-2 customCheckbox"
                                                        checked={formik.values.employerKycApproved}
                                                        onChange={handleExclusiveCheckboxChange("employerKycApproved", [
                                                            "employerKycPending",
                                                            "employerKycUploadedNotApproved",
                                                            "employerKycApproved",
                                                        ])}
                                                    />
                                                </div>
                                            </div>

                                            <div className="mb-3">
                                                <strong>C.</strong>{" "}
                                                In case the person was earlier a member of EPF
                                                Scheme, 1952 and EPS, 1995:
                                                <ul className="mt-2 mb-2">
                                                    <li>
                                                        The above member ID of the member as mentioned
                                                        in (A) above has been tagged with his/her
                                                        UAN/Previous Member ID as declared by member.
                                                    </li>

                                                    <li className="mt-2">
                                                        <strong>Please tick the appropriate option:</strong>
                                                    </li>
                                                </ul>

                                                <div className="ms-4">
                                                    <Form.Check
                                                        name="employerTransferApproved"
                                                        type="checkbox"
                                                        label={
                                                            <span
                                                                style={{
                                                                    marginLeft: "6px",
                                                                }}
                                                            >
                                                                The KYC details of the above member in the UAN database have been approved with Digital Signature Certificate and transfer request has been generated on portal.
                                                            </span>}
                                                        className="mb-2 d-block customCheckbox"
                                                        checked={formik.values.employerTransferApproved}
                                                        onChange={handleExclusiveCheckboxChange("employerTransferApproved", [
                                                            "employerTransferApproved",
                                                            "employerPhysicalClaim",
                                                        ])}
                                                    />

                                                    <Form.Check
                                                        name="employerPhysicalClaim"
                                                        type="checkbox"
                                                        label={
                                                            <span
                                                                style={{
                                                                    marginLeft: "6px",
                                                                }}
                                                            >
                                                                As the DSC of establishment are not registered with EPFO, the member has been informed to file physical claim (Form-13) for transfer of funds from his previous establishment.
                                                            </span>}
                                                        className="mb-2 customCheckbox"
                                                        checked={formik.values.employerPhysicalClaim}
                                                        onChange={handleExclusiveCheckboxChange("employerPhysicalClaim", [
                                                            "employerTransferApproved",
                                                            "employerPhysicalClaim",
                                                        ])}
                                                    />
                                                </div>
                                            </div>

                                            <Row className="mt-5 align-items-end">
                                                <Col md={6}>
                                                    <strong>DATE:</strong>
                                                    <div className="mt-2" style={{ maxWidth: "260px" }}>
                                                        <DatePickerInput
                                                            name="employerDate"
                                                            value={formik.values.employerDate}
                                                            onValueChange={createDateValueChangeHandler(formik.setFieldValue, "employerDate")}
                                                        />
                                                    </div>
                                                </Col>

                                                <Col md={6} className="text-center">
                                                    <div
                                                        style={{
                                                            fontFamily: "cursive",
                                                            fontSize: "34px",
                                                            transform: "rotate(-4deg)",
                                                            color: "#111"
                                                        }}
                                                    >
                                                        B. N. N Murthy
                                                    </div>

                                                    <div className="fw-bold mt-2 text-uppercase">
                                                        Signature of Employer with Seal of
                                                        Establishment
                                                    </div>
                                                </Col>
                                            </Row>
                                        </Card.Body>
                                    </Card>
                                )}
                            </Card.Body>
                        </Card>

                        {isSubmitted && (
                            <Alert variant="success" className="mt-4 mb-0">
                                All sections have been completed successfully.
                            </Alert>
                        )}

                        <div className="text-end mt-4">
                            <Button
                                type="button"
                                variant="secondary"
                                className="me-2"
                                onClick={() => formik.resetForm()}
                            >
                                Clear
                            </Button>
                            <Button
                                type="submit"
                                style={{ background: "#f18200", border: "none" }}
                                disabled={formik.isSubmitting || !!isSubmitted}
                            >
                                {formik.isSubmitting ? "Submitting..." : "Submit"}
                            </Button>
                            {employeePFData?.EmailID !== 'hr@natit.in' && (
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

        </Container>
    );
}
