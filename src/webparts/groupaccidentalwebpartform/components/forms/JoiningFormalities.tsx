import * as React from 'react';
import FormSection from './FormSection';
import { Form, Row, Col, Table, Card, Button } from 'react-bootstrap';
import { useFormik } from 'formik';
import * as Yup from 'yup';
import {
    createDateValueChangeHandler,
    createMatchingDateValidation,
    createDateValidation,
    DatePickerInput,
    formatDateInput,
    formatSharePointDateForInput,
    parseDobForSharePoint,
    parseFormDateForSharePoint
} from './dateFieldUtils';
import SignatureUpload from './SignatureUpload';
import { getSP } from '../../../../pnpjsConfig';
import type { ISequentialFormProps } from './ISequentialFormProps';
import moment from 'moment';
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import {
    candidateFieldNames,
    getCandidateDateValue,
    getCandidateValue,
    setFieldIfEmpty
} from './candidateAutoFillUtils';
import { replaceListItemAttachment, upsertByCanId } from './sharePointListUtils';

type EducationRow = {
    qualification: string;
    institute: string;
    specialization: string;
    year: string;
};
type Reference = {
    name: string;
    occupation: string;
    relationship: string;
    contact: string;
};
const maxPhotoSizeInBytes = 2 * 1024 * 1024;

const toAttachmentPreviewUrl = (serverRelativeUrl?: string): string => {
    if (!serverRelativeUrl) {
        return '';
    }
    if (serverRelativeUrl.startsWith('http')) {
        return serverRelativeUrl;
    }
    return `${window.location.origin}${serverRelativeUrl}`;
};
const JoiningFormalities = ({
    onComplete,
    sharedEmployeeSignature,
    onEmployeeSignatureChange,
    context,
    employeePFData,
    sharedDateOfBirth,
    candidateId,
    isAdminEdit,
    isFinallySubmitted,
    submitButtonLabel = 'Continue',
    onSharedDateOfBirthChange,
}: ISequentialFormProps): JSX.Element => {
    const readOnly = !!isFinallySubmitted && !isAdminEdit;
    const [getDepartments, setGetDepartments] = React.useState<any[]>([]);
    const [getDesignations, setGetDesignations] = React.useState<any[]>([]);
    const [getEducationsDetails, setGetEducationsDetails] = React.useState<any[]>([]);
    const formRef = React.useRef<HTMLDivElement>(null);

    React.useEffect(() => {
        if (!document.getElementById('bootstrap-css')) {
            const link = document.createElement('link');
            link.id = 'bootstrap-css';
            link.rel = 'stylesheet';
            link.href = 'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css';
            document.head.appendChild(link);
        }
    }, []);
    const commonDateOfBirth =
        sharedDateOfBirth || getCandidateDateValue(employeePFData, candidateFieldNames.dateOfBirth);
    const commonDateOfBirthMessage = commonDateOfBirth
        ? `Date of Birth must match the employee Date of Birth (${commonDateOfBirth})`
        : "Date of Birth must match the employee Date of Birth";

    const joiningValidationSchema = Yup.object().shape({
        employeeId: Yup.string(),
        designation: Yup.string(),
        reportingTo: Yup.string(),
        department: Yup.string(),
        fullName: Yup.string().required("Full Name is required"),
        dob: createMatchingDateValidation(
            createDateValidation("Date of Birth is required", "Date of Birth must be in DD/MM/YYYY format"),
            commonDateOfBirth,
            commonDateOfBirthMessage
        )
            .test(
                "dob-age-validation",
                "Minimum age must be 18 years",
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
        actualDob: createDateValidation("Actual DOB is required", "Actual DOB must be in DD/MM/YYYY format")
            .test(
                "actualDob-age-validation",
                "Minimum age must be 18 years",
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
        panNo: Yup.string().matches(
            /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/,
            {
                message: "Invalid PAN Number",
                excludeEmptyString: true
            }
        ),
        photoFile: Yup.mixed<File>()
            .required("Photo file is required")
            .test(
                "fileSize",
                "Photo file size must be below 2 MB",
                (value) => !value || value.size <= maxPhotoSizeInBytes
            ),
        education: Yup.array().of(
            Yup.object().shape({
                qualification: Yup.string().test(
                    "first-qualification-required",
                    "Qualification is required",
                    function (value) {
                        const index = Number(this.path.match(/\[(\d+)\]/)?.[1] ?? -1);
                        return index !== 0 || !!value?.trim();
                    }
                ),
                institute: Yup.string().test(
                    "first-institute-required",
                    "Institute is required",
                    function (value) {
                        const index = Number(this.path.match(/\[(\d+)\]/)?.[1] ?? -1);
                        return index !== 0 || !!value?.trim();
                    }
                ),
                specialization: Yup.string().test(
                    "first-specialization-required",
                    "Specialization is required",
                    function (value) {
                        const index = Number(this.path.match(/\[(\d+)\]/)?.[1] ?? -1);
                        return index !== 0 || !!value?.trim();
                    }
                ),
                year: Yup.string().test(
                    "first-year-required",
                    "Year is required",
                    function (value) {
                        const index = Number(this.path.match(/\[(\d+)\]/)?.[1] ?? -1);
                        return index !== 0 || !!value?.trim();
                    }
                ).matches(/^\d+$/, "Year must contain only numeric values"),
            })
        ),
        /* ================= BANK ================= */
        bankName: Yup.string().required("Bank name is required"),
        accountNo: Yup.string()
            .required("Account number is required")
            .matches(/^[0-9]{9,18}$/, "Invalid account number"),
        ifscCode: Yup.string()
            .required("IFSC is required")
            .matches(/^[A-Z]{4}0[A-Z0-9]{6}$/, "Invalid IFSC"),
        branchDetails: Yup.string().required("Branch details required"),
        /* ================= REFERENCES ================= */
        references: Yup.array().of(
            Yup.object().shape({
                name: Yup.string().required("Name is required"),
                occupation: Yup.string().required("Occupation is required"),
                relationship: Yup.string().required("Relationship is required"),
                contact: Yup.string()
                    .required("Contact No & Address is required"),
            })
        ),
        /* ================= JOINING LETTER ================= */
        joiningLetterDate: createDateValidation("Date is required", "Date must be in DD/MM/YYYY format"),
        joiningDateText: createDateValidation("Joining date required", "Joining date must be in DD/MM/YYYY format"),
        designationText: Yup.string().required("Designation required"),
        signatureName: Yup.mixed()
            .required("Signature required")
            .test(
                "fileSize",
                "Signature should be less than 200 KB",
                (value: any) => {
                    if (!value) return false;

                    // when value comes from sharedEmployeeSignature
                    if (typeof value === "string") {
                        return true;
                    }

                    // uploaded file validation
                    return value.size <= 200 * 1024;
                }
            )
    });
    const joiningFormValidation = useFormik<{
        employeeId: string;
        designation: string;
        reportingTo: string;
        department: string;
        fullName: string;
        dob: string;
        actualDob: string;
        panNo: string;
        presentAddress: string;
        permanentAddress: string;
        fatherName: string;
        maritalStatus: string;
        spouseName: string;
        photoFile: File | string | null;
        education: EducationRow[];
        bankName: string;
        accountNo: string;
        ifscCode: string;
        branchDetails: string;
        references: Reference[];
        joiningLetterDate: string;
        joiningDateText: string;
        designationText: string;
        signatureName: File | string;
    }>({
        initialValues: {
            employeeId: "",
            designation: "",
            reportingTo: "",
            department: "",
            fullName: "",
            dob: "",
            actualDob: "",
            panNo: "",
            presentAddress: "",
            permanentAddress: "",
            fatherName: "",
            maritalStatus: "",
            spouseName: "",
            photoFile: null,
            education: [
                {
                    qualification: "",
                    institute: "",
                    specialization: "",
                    year: "",
                },
                {
                    qualification: "",
                    institute: "",
                    specialization: "",
                    year: "",
                },
                {
                    qualification: "",
                    institute: "",
                    specialization: "",
                    year: "",
                },
            ],
            bankName: "",
            accountNo: "",
            ifscCode: "",
            branchDetails: "",
            references: [
                { name: "", occupation: "", relationship: "", contact: "" },
                { name: "", occupation: "", relationship: "", contact: "" },
            ],
            joiningLetterDate: "",
            joiningDateText: "",
            designationText: "",
            signatureName: "",
        },
        validationSchema: joiningValidationSchema,
        onSubmit: async (values) => {
            try {
                const sp = getSP(context);
                if (!values.fullName) return;

                const employeeDob = parseDobForSharePoint(values.dob);
                const employeeActualDob = parseDobForSharePoint(values.actualDob);
                const joiningLetterDate = parseFormDateForSharePoint(
                    values.joiningLetterDate,
                    'datetime'
                );
                const joiningDateText = parseFormDateForSharePoint(
                    values.joiningDateText,
                    'datetime'
                );

                if (!employeeDob || !employeeActualDob || !joiningLetterDate || !joiningDateText) {
                    throw new Error(
                        'Invalid date on the form. Please enter all dates in DD/MM/YYYY format.'
                    );
                }

                const employeeSPID = await upsertByCanId(
                    sp,
                    'JoiningFormalities',
                    String(employeePFData?.ID),
                    {
                        Title: values.fullName,
                        hr_employee_ID: values.employeeId,
                        hr_designation: values.designation,
                        hr_department: values.department,
                        employee_full_name: values.fullName,
                        employee_DOB: employeeDob,
                        employee_actual_DOB: employeeActualDob,
                        employee_reporting_to: values.reportingTo,
                        bank_name_as_per_bank_records: values.bankName,
                        bank_account_no: values.accountNo,
                        bank_IFSC_code: values.ifscCode,
                        bank_branch: values.branchDetails,
                        joining_letter_date: joiningLetterDate,
                        joining_date_text: joiningDateText,
                        designation_text: values.designationText
                    }
                );

                if (values.photoFile && typeof values.photoFile !== 'string') {
                    await replaceListItemAttachment(
                        sp,
                        'JoiningFormalities',
                        employeeSPID,
                        values.photoFile.name,
                        values.photoFile
                    );
                }

                if (values.signatureName && typeof values.signatureName !== 'string') {
                    await replaceListItemAttachment(
                        sp,
                        'JoiningFormalities',
                        employeeSPID,
                        values.signatureName.name,
                        values.signatureName
                    );
                }

                const existingEducation = await sp.web.lists
                    .getByTitle('EmployeeEducation')
                    .items.filter(`EmployeeID eq '${employeeSPID}'`)();
                for (const row of existingEducation) {
                    await sp.web.lists
                        .getByTitle('EmployeeEducation')
                        .items.getById(row.Id)
                        .delete();
                }
                const existingReferences = await sp.web.lists
                    .getByTitle('EmployeeReferences')
                    .items.filter(`EmployeeID eq '${employeeSPID}'`)();
                for (const row of existingReferences) {
                    await sp.web.lists
                        .getByTitle('EmployeeReferences')
                        .items.getById(row.Id)
                        .delete();
                }

                if (Array.isArray(values.education)) {
                    for (const edu of values.education) {
                        if (!edu.qualification) continue;
                        await sp.web.lists.getByTitle("EmployeeEducation").items.add({
                            Title: values.fullName,
                            EmployeeID: String(employeeSPID),
                            Qualification: edu.qualification,
                            Institute: edu.institute,
                            Specialization: edu.specialization,
                            YearCompleted: edu.year?.trim()
                                ? parseFormDateForSharePoint(`01/01/${edu.year.trim()}`, 'datetime')
                                : null
                        });
                    }
                }
                if (Array.isArray(values.references)) {
                    for (const ref of values.references) {
                        if (!ref.name) continue;
                        await sp.web.lists.getByTitle("EmployeeReferences").items.add({
                            Title: values.fullName,
                            EmployeeID: String(employeeSPID),
                            Name: ref.name,
                            Occupation: ref.occupation,
                            Relationship: ref.relationship,
                            Contact: ref.contact
                        });
                    }
                }
                onComplete?.();
            } catch (error) {
                alert(
                    error instanceof Error
                        ? error.message
                        : 'Submission failed. Please check all dates are in DD/MM/YYYY format.'
                );
            }
        }
    });
    React.useEffect(() => {
        const autoFill = async (): Promise<void> => {
            const designation = getCandidateValue(employeePFData, candidateFieldNames.designation);
            const joiningDate = getCandidateDateValue(employeePFData, candidateFieldNames.joiningDate);

            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'employeeId', getCandidateValue(employeePFData, candidateFieldNames.employeeId));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'designation', designation);
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'department', getCandidateValue(employeePFData, candidateFieldNames.department));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'fullName', getCandidateValue(employeePFData, candidateFieldNames.employeeName));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'dob', getCandidateDateValue(employeePFData, candidateFieldNames.dateOfBirth));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'actualDob', getCandidateDateValue(employeePFData, candidateFieldNames.dateOfBirth));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'panNo', getCandidateValue(employeePFData, candidateFieldNames.panNo));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'presentAddress', getCandidateValue(employeePFData, candidateFieldNames.address));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'permanentAddress', getCandidateValue(employeePFData, candidateFieldNames.address));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'fatherName', getCandidateValue(employeePFData, candidateFieldNames.fatherOrHusbandName));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'maritalStatus', getCandidateValue(employeePFData, candidateFieldNames.maritalStatus));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'spouseName', getCandidateValue(employeePFData, candidateFieldNames.spouseName));
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'joiningDateText', joiningDate);
            await setFieldIfEmpty(joiningFormValidation.values, joiningFormValidation.setFieldValue, 'designationText', designation);
        };

        void autoFill();
    }, [employeePFData]);
    React.useEffect(() => {

        const loadSavedFormData = async (): Promise<void> => {

            try {

                const sp = getSP(context);

                if (!employeePFData?.ID) return;

                const items = await sp.web.lists
                    .getByTitle("JoiningFormalities")
                    .items
                    .filter(`can_id eq '${employeePFData.ID}'`)
                    .top(1)
                    .orderBy("ID", false)();

                if (items.length === 0) return;

                const data = items[0];

                const employeeSPID = data.ID;

                // EDUCATION
                const educationItems = await sp.web.lists
                    .getByTitle("EmployeeEducation")
                    .items
                    .filter(`EmployeeID eq '${employeeSPID}'`)();

                // REFERENCES
                const referenceItems = await sp.web.lists
                    .getByTitle("EmployeeReferences")
                    .items
                    .filter(`EmployeeID eq '${employeeSPID}'`)();

                // ATTACHMENTS
                const attachments = await sp.web.lists
                    .getByTitle("JoiningFormalities")
                    .items.getById(employeeSPID)
                    .attachmentFiles();

                const imageAttachments = attachments.filter((a: { FileName?: string }) =>
                    !!a.FileName?.match(/\.(jpg|jpeg|png|gif|webp)$/i)
                );

                const signatureAttachment =
                    imageAttachments.find((a: { FileName?: string }) =>
                        /sign/i.test(a.FileName || '')
                    ) ??
                    (imageAttachments.length > 1
                        ? imageAttachments[imageAttachments.length - 1]
                        : undefined);

                const photoAttachment = imageAttachments.find(
                    (a: { FileName?: string; ServerRelativeUrl?: string }) =>
                        a !== signatureAttachment
                );

                const photoUrl = toAttachmentPreviewUrl(photoAttachment?.ServerRelativeUrl);
                const signatureUrl = toAttachmentPreviewUrl(
                    signatureAttachment?.ServerRelativeUrl
                );

                await joiningFormValidation.setValues({

                    ...joiningFormValidation.values,

                    employeeId: data.hr_employee_ID || "",

                    designation: data.hr_designation || "",

                    department: data.hr_department || "",

                    reportingTo: data.employee_reporting_to || "",

                    fullName: data.employee_full_name || "",

                    dob: formatSharePointDateForInput(data.employee_DOB as string),

                    actualDob: formatSharePointDateForInput(data.employee_actual_DOB as string),

                    bankName: data.bank_name_as_per_bank_records || "",

                    accountNo: data.bank_account_no || "",

                    ifscCode: data.bank_IFSC_code || "",

                    branchDetails: data.bank_branch || "",

                    photoFile: photoUrl || null,

                    education:
                        educationItems.length > 0
                            ? educationItems.map((edu: any) => ({
                                qualification: edu.Qualification || "",
                                institute: edu.Institute || "",
                                specialization: edu.Specialization || "",
                                year: edu.YearCompleted
                                    ? formatSharePointDateForInput(edu.YearCompleted as string).split('/')[2] ||
                                      ''
                                    : ""
                            }))
                            : [],

                    joiningLetterDate: formatSharePointDateForInput(
                        data.joining_letter_date as string
                    ),

                    joiningDateText: formatSharePointDateForInput(
                        data.joining_date_text as string
                    ),

                    designationText: data.designation_text || "",

                    signatureName: signatureUrl || joiningFormValidation.values.signatureName || "",
                    references:
                        referenceItems.length > 0
                            ? referenceItems.map((ref: any) => ({
                                name: ref.Name || "",
                                occupation: ref.Occupation || "",
                                relationship: ref.Relationship || "",
                                contact: ref.Contact || ""
                            }))
                            : []

                });

                const loadedDob = formatSharePointDateForInput(data.employee_DOB as string);
                if (loadedDob) {
                    onSharedDateOfBirthChange?.(loadedDob);
                }

                if (signatureUrl) {
                    onEmployeeSignatureChange?.(signatureUrl);
                }

            } catch {
                // Saved form data could not be loaded.
            }

        };

        void loadSavedFormData();

    }, [employeePFData]);
    React.useEffect(() => {
        if (sharedEmployeeSignature && joiningFormValidation.values.signatureName !== sharedEmployeeSignature) {
            joiningFormValidation.setFieldValue('signatureName', sharedEmployeeSignature).catch(() => undefined);
        }
    }, [sharedEmployeeSignature]);
    const handleDateOfBirthChange = (value: string): void => {
        const formattedValue = formatDateInput(value);
        joiningFormValidation.setFieldValue('dob', formattedValue).catch(() => undefined);
        joiningFormValidation.setFieldValue('actualDob', formattedValue).catch(() => undefined);
        onSharedDateOfBirthChange?.(formattedValue);
    };

    const getDepart = async (): Promise<void> => {
        try {
            const sp = getSP(context);
            const items = await sp.web.lists
                .getByTitle("DepartmentsList")
                .items();
            setGetDepartments(items);
        } catch {
            // Departments list unavailable.
        }
    };
    const getDesi = async (): Promise<void> => {
        try {
            const sp = getSP(context);
            const items = await sp.web.lists
                .getByTitle("DesignationsList")
                .items();
            setGetDesignations(items);
        } catch {
            // Designations list unavailable.
        }
    };
    const getEducations = async (): Promise<void> => {
        try {
            const sp = getSP(context);
            const items = await sp.web.lists
                .getByTitle("EducationalBackground")
                .items();
            setGetEducationsDetails(items);
        } catch {
            // Education options list unavailable.
        }
    };
    React.useEffect(() => {
        void getDepart();
        void getDesi();
        void getEducations();
    }, []);

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

        pdf.save("Joining-Formalities.pdf");
    };

    return (
        <>
            <FormSection title="Joining Formalities" />
            <div>
                <div ref={formRef} className="no-break">
                    <Card className="shadow-lg border-0">
                        <Card.Body>
                            <Form onSubmit={joiningFormValidation.handleSubmit}>
                                <div className="align-items-start mb-4 row">
                                    <div className="col-md-8">
                                        <div className="text-center text-md-start">
                                            <h3 className="fw-bold">EMPLOYEE INFORMATION SHEET testing</h3>
                                            <p className="mb-0">NAT IT Services Pvt Ltd</p>
                                            <small>Plot No:21, Sruthi Sadan, Gachibowli, Hyderabad</small>
                                        </div>
                                    </div>
                                    <div className="d-flex justify-content-md-end justify-content-center mt-3 mt-md-0 col-md-4">
                                        <div style={{ width: "200px" }}>
                                            <div
                                                style={{
                                                    position: "relative",
                                                    width: "200px",
                                                    height: "200px",
                                                }}
                                            >
                                                {/* INPUT */}
                                                <input
                                                    type="file"
                                                    name="photoFile"
                                                    accept="image/*"
                                                    onChange={(e) => {
                                                        const file = e.currentTarget.files?.[0] ?? null;
                                                        joiningFormValidation
                                                            .setFieldValue("photoFile", file)
                                                            .catch(() => undefined);
                                                    }}
                                                    onBlur={() => {
                                                        joiningFormValidation
                                                            .setFieldTouched("photoFile", true)
                                                            .catch(() => undefined);
                                                    }}
                                                    style={{
                                                        position: "absolute",
                                                        width: "100%",
                                                        height: "100%",
                                                        opacity: 0,
                                                        cursor: "pointer",
                                                        zIndex: 2,
                                                    }}
                                                />
                                                {/* UPLOAD BOX */}
                                                <div
                                                    style={{
                                                        width: "200px",
                                                        height: "200px",
                                                        border: "2px dashed #6c757d",
                                                        borderRadius: "10px",
                                                        display: "flex",
                                                        alignItems: "center",
                                                        justifyContent: "center",
                                                        textAlign: "center",
                                                        padding: "10px",
                                                        background: "#f8f9fa",
                                                        cursor: "pointer",
                                                    }}
                                                >
                                                    {joiningFormValidation.values.photoFile ? (
                                                        <img
                                                            src={
                                                                typeof joiningFormValidation.values.photoFile === "string"
                                                                    ? joiningFormValidation.values.photoFile
                                                                    : joiningFormValidation.values.photoFile
                                                                        ? URL.createObjectURL(joiningFormValidation.values.photoFile)
                                                                        : ""
                                                            }
                                                            alt="Employee preview"
                                                            style={{
                                                                maxWidth: "100%",
                                                                maxHeight: "100%",
                                                                objectFit: "cover",
                                                                borderRadius: "8px",
                                                            }}
                                                        />
                                                    ) : (
                                                        <>
                                                            Click to Upload <br />
                                                            Photo
                                                        </>
                                                    )}
                                                </div>
                                            </div>
                                            {/* FILE NAME */}
                                            {/* {joiningFormValidation.values.photoFile && (
                                                <small className="d-block mt-2 text-center text-muted">
                                                    {joiningFormValidation.values.photoFile.name}
                                                </small>
                                            )}
                                            {joiningFormValidation.touched.photoFile &&
                                                joiningFormValidation.errors.photoFile && (
                                                    <p className="mt-2 mb-0 text-center text-danger small">
                                                        {joiningFormValidation.errors.photoFile}
                                                    </p>
                                                )} */}
                                            {joiningFormValidation.values.photoFile && (
                                                <small className="d-block mt-2 text-center text-muted">

                                                    {typeof joiningFormValidation.values.photoFile === "string"
                                                        ? joiningFormValidation.values.photoFile.split("/").pop()
                                                        : joiningFormValidation.values.photoFile.name}

                                                </small>
                                            )}

                                            {joiningFormValidation.touched.photoFile &&
                                                joiningFormValidation.errors.photoFile && (
                                                    <p className="mt-2 mb-0 text-center text-danger small">
                                                        {joiningFormValidation.errors.photoFile}
                                                    </p>
                                                )}
                                        </div>
                                    </div>
                                </div>
                                {/* Employee Details */}
                                {isAdminEdit && (
                                    <div>
                                        <h5
                                            className="text-white p-2 rounded"
                                            style={{ backgroundColor: "#f18200" }}
                                        >
                                            Employee Details
                                        </h5>
                                        <Row className="mb-3">
                                            <Col md={6}>
                                                <Form.Label>Employee ID</Form.Label>
                                                <Form.Control
                                                    name="employeeId"
                                                    value={joiningFormValidation.values.employeeId}
                                                    onChange={joiningFormValidation.handleChange}
                                                    onBlur={joiningFormValidation.handleBlur}

                                                />
                                                {joiningFormValidation.touched.employeeId &&
                                                    joiningFormValidation.errors.employeeId && (
                                                        <p className="text-danger small">
                                                            {joiningFormValidation.errors.employeeId}
                                                        </p>
                                                    )}
                                            </Col>
                                            <Col md={6}>
                                                <Form.Label>Designation</Form.Label>

                                                <Form.Select
                                                    name="designation"
                                                    value={joiningFormValidation.values.designation}
                                                    onChange={joiningFormValidation.handleChange}
                                                    onBlur={joiningFormValidation.handleBlur}

                                                >
                                                    <option value="">Select Designation</option>
                                                    {getDesignations?.map((desig) => (
                                                        <option key={desig.ID} value={desig.Designations}>
                                                            {desig.Designations}
                                                        </option>
                                                    ))}
                                                </Form.Select>
                                                {joiningFormValidation.touched.designation &&
                                                    joiningFormValidation.errors.designation && (
                                                        <p className="text-danger small">
                                                            {joiningFormValidation.errors.designation}
                                                        </p>
                                                    )}
                                            </Col>
                                        </Row>
                                        <Row className="mb-3">
                                            <Col md={6}>
                                                <Form.Label>Reporting To</Form.Label>
                                                <Form.Control
                                                    name="reportingTo"
                                                    value={joiningFormValidation.values.reportingTo}
                                                    onChange={joiningFormValidation.handleChange}
                                                    onBlur={joiningFormValidation.handleBlur}

                                                />
                                                {joiningFormValidation.touched.reportingTo &&
                                                    joiningFormValidation.errors.reportingTo && (
                                                        <p className="text-danger small">
                                                            {joiningFormValidation.errors.reportingTo}
                                                        </p>
                                                    )}
                                            </Col>
                                            <Col md={6}>
                                                <Form.Label>Department</Form.Label>

                                                <Form.Select
                                                    name="department"
                                                    value={joiningFormValidation.values.department}
                                                    onChange={joiningFormValidation.handleChange}
                                                    onBlur={joiningFormValidation.handleBlur}

                                                >
                                                    <option value="">Select Department</option>
                                                    {getDepartments?.map((dept) => (
                                                        <option key={dept.ID} value={dept.Department}>
                                                            {dept.Department}
                                                        </option>
                                                    ))}
                                                </Form.Select>
                                                {joiningFormValidation.touched.department &&
                                                    joiningFormValidation.errors.department && (
                                                        <p className="text-danger small">
                                                            {joiningFormValidation.errors.department}
                                                        </p>
                                                    )}
                                            </Col>
                                        </Row>
                                    </div>
                                )}

                                {/* Personal Info */}
                                <h5
                                    className="text-white p-2 rounded"
                                    style={{ backgroundColor: "#f18200" }}
                                >
                                    Personal Information
                                </h5>
                                <Row className="mb-3">
                                    <Col md={6}>
                                        <Form.Label>Full Name</Form.Label>
                                        <Form.Control
                                            name="fullName"
                                            value={joiningFormValidation.values.fullName}
                                            onChange={joiningFormValidation.handleChange}
                                            onBlur={joiningFormValidation.handleBlur}
                                        />
                                        {joiningFormValidation.touched.fullName &&
                                            joiningFormValidation.errors.fullName && (
                                                <p className="text-danger small">
                                                    {joiningFormValidation.errors.fullName}
                                                </p>
                                            )}
                                    </Col>
                                    <Col md={3}>
                                        <Form.Label>DOB</Form.Label>
                                        <DatePickerInput
                                            name="dob"
                                            value={joiningFormValidation.values.dob}
                                            onValueChange={handleDateOfBirthChange}
                                            onBlur={joiningFormValidation.handleBlur}
                                        />
                                        {joiningFormValidation.touched.dob &&
                                            joiningFormValidation.errors.dob && (
                                                <p className="text-danger small">
                                                    {joiningFormValidation.errors.dob}
                                                </p>
                                            )}
                                    </Col>
                                    <Col md={3}>
                                        <Form.Label>Actual DOB</Form.Label>
                                        <DatePickerInput
                                            name="actualDob"
                                            value={joiningFormValidation.values.actualDob}
                                            onValueChange={createDateValueChangeHandler(joiningFormValidation.setFieldValue, "actualDob")}
                                            onBlur={joiningFormValidation.handleBlur}
                                        />
                                        {joiningFormValidation.touched.actualDob &&
                                            joiningFormValidation.errors.actualDob && (
                                                <p className="text-danger small">
                                                    {joiningFormValidation.errors.actualDob}
                                                </p>
                                            )}
                                    </Col>
                                </Row>
                                {/* Education */}
                                <h5
                                    className="text-white p-2 rounded"
                                    style={{ backgroundColor: "#f18200" }}
                                >
                                    Educational Background
                                </h5>
                                <Table bordered responsive>
                                    <thead>
                                        <tr>
                                            <th>Qualification</th>
                                            <th>Institute</th>
                                            <th>Specialization</th>
                                            <th>Year</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {joiningFormValidation.values.education.map((row, index) => {
                                            const touchedRow = joiningFormValidation.touched.education?.[index];
                                            const errorRow = joiningFormValidation.errors.education?.[index];
                                            const isFirstRow = index === 0;
                                            return (
                                                <tr key={index}>
                                                    <td>
                                                        {/* <Form.Control
                                                            
                                                            
                                                            
                                                            
                                                            
                                                        /> */}
                                                        <Form.Select
                                                            name={`education[${index}].qualification`}
                                                            value={row.qualification}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            isInvalid={isFirstRow && typeof touchedRow === "object" && touchedRow?.qualification && typeof errorRow === "object" && !!errorRow?.qualification}
                                                        >
                                                            <option value="">Select Qualification</option>
                                                            {getEducationsDetails?.map((edu) => (
                                                                <option key={edu.ID} value={edu.qualification}>
                                                                    {edu.qualification}
                                                                </option>
                                                            ))}
                                                        </Form.Select>
                                                        {isFirstRow && typeof touchedRow === "object" && touchedRow?.qualification && typeof errorRow === "object" && errorRow?.qualification && (
                                                            <p className="text-danger small">
                                                                {errorRow.qualification}
                                                            </p>
                                                        )}
                                                    </td>
                                                    <td>
                                                        <Form.Control
                                                            name={`education[${index}].institute`}
                                                            value={row.institute}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            isInvalid={isFirstRow && typeof touchedRow === "object" && touchedRow?.institute && typeof errorRow === "object" && !!errorRow?.institute}
                                                        />
                                                        {isFirstRow && typeof touchedRow === "object" && touchedRow?.institute && typeof errorRow === "object" && errorRow?.institute && (
                                                            <p className="text-danger small">
                                                                {errorRow.institute}
                                                            </p>
                                                        )}
                                                    </td>
                                                    <td>
                                                        <Form.Control
                                                            name={`education[${index}].specialization`}
                                                            value={row.specialization}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            isInvalid={isFirstRow && typeof touchedRow === "object" && touchedRow?.specialization && typeof errorRow === "object" && !!errorRow?.specialization}
                                                        />
                                                        {isFirstRow && typeof touchedRow === "object" && touchedRow?.specialization && typeof errorRow === "object" && errorRow?.specialization && (
                                                            <p className="text-danger small">
                                                                {errorRow.specialization}
                                                            </p>
                                                        )}
                                                    </td>
                                                    <td>
                                                        <Form.Control
                                                            name={`education[${index}].year`}
                                                            value={row.year}
                                                            maxLength={4}
                                                            onChange={(e) => {
                                                                const value = e.target.value.replace(/[^0-9]/g, '');
                                                                joiningFormValidation.setFieldValue(`education[${index}].year`, value).catch(() => undefined);
                                                            }}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            isInvalid={isFirstRow && typeof touchedRow === "object" && touchedRow?.year && typeof errorRow === "object" && !!errorRow?.year}
                                                            placeholder="YYYY"
                                                        />
                                                        {isFirstRow && typeof touchedRow === "object" && touchedRow?.year && typeof errorRow === "object" && errorRow?.year && (
                                                            <p className="text-danger small">
                                                                {errorRow.year}
                                                            </p>
                                                        )}
                                                    </td>
                                                </tr>
                                            );
                                        })}
                                    </tbody>
                                </Table>
                                <div>
                                    <h5 className="text-dark py-2">Bank Account Details</h5>
                                    <div className="table-responsive">
                                        <table className="table table-bordered">
                                            <thead>
                                                <tr>
                                                    <th> Bank Name</th>
                                                    <th>Name as per Bank Records</th>
                                                    <th>Account No</th>
                                                    <th>IFSC Code</th>
                                                    <th>Branch Details</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <tr>
                                                    {/* Bank Name */}
                                                    <td>
                                                        <select disabled className='form-select'>
                                                            <option value="ICICI Bank">ICICI Bank</option>
                                                            <option value="Hdfc Bank">HDFC Bank</option>
                                                        </select>
                                                    </td>
                                                    <td>
                                                        <input
                                                            name="bankName"
                                                            className="form-control"
                                                            value={joiningFormValidation.values.bankName}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                        />
                                                        {joiningFormValidation.touched.bankName &&
                                                            joiningFormValidation.errors.bankName && (
                                                                <p className="text-danger small">
                                                                    {joiningFormValidation.errors.bankName}
                                                                </p>
                                                            )}
                                                    </td>
                                                    {/* Account No */}
                                                    <td>
                                                        <input
                                                            name="accountNo"
                                                            className="form-control"
                                                            value={joiningFormValidation.values.accountNo}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                        />
                                                        {joiningFormValidation.touched.accountNo &&
                                                            joiningFormValidation.errors.accountNo && (
                                                                <p className="text-danger small">
                                                                    {joiningFormValidation.errors.accountNo}
                                                                </p>
                                                            )}
                                                    </td>
                                                    {/* IFSC */}
                                                    <td>
                                                        <input
                                                            name="ifscCode"
                                                            className="form-control"
                                                            value={joiningFormValidation.values.ifscCode}
                                                            onChange={(e) => {
                                                                joiningFormValidation
                                                                    .setFieldValue("ifscCode", e.target.value.toUpperCase())
                                                                    .catch(() => undefined);
                                                            }}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                        />
                                                        {joiningFormValidation.touched.ifscCode &&
                                                            joiningFormValidation.errors.ifscCode && (
                                                                <p className="text-danger small">
                                                                    {joiningFormValidation.errors.ifscCode}
                                                                </p>
                                                            )}
                                                    </td>
                                                    {/* Branch */}
                                                    <td>
                                                        <input
                                                            name="branchDetails"
                                                            className="form-control"
                                                            value={joiningFormValidation.values.branchDetails}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                        />
                                                        {joiningFormValidation.touched.branchDetails &&
                                                            joiningFormValidation.errors.branchDetails && (
                                                                <p className="text-danger small">
                                                                    {joiningFormValidation.errors.branchDetails}
                                                                </p>
                                                            )}
                                                    </td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                                <div className="g-4 mt-2 row">
                                    <h5 className="text-dark py-2">References</h5>
                                    <h5
                                        className="text-white p-2 mt-2 rounded"
                                        style={{ backgroundColor: "#f18200" }}
                                    >
                                        REFERENCES : (Kindly provide two references e.g. Reporting Managers)
                                    </h5>
                                    {joiningFormValidation.values.references.map((ref, index) => {
                                        const touchedReference = joiningFormValidation.touched.references?.[index];
                                        const errorReference = joiningFormValidation.errors.references?.[index];
                                        return (
                                            <div className="col-md-6" key={index}>
                                                <div className="border rounded p-3 shadow-sm">
                                                    <h6 className="fw-bold mb-3">
                                                        Reference {index + 1}
                                                    </h6>
                                                    {/* Name */}
                                                    <div className="mb-3">
                                                        <label className="form-label">Name</label>
                                                        <input
                                                            name={`references[${index}].name`}
                                                            value={ref.name}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            className="form-control"
                                                        />
                                                        {typeof touchedReference === "object" &&
                                                            touchedReference?.name &&
                                                            typeof errorReference === "object" &&
                                                            errorReference?.name && (
                                                                <p className="text-danger small">
                                                                    {errorReference.name}
                                                                </p>
                                                            )}
                                                    </div>
                                                    {/* Occupation */}
                                                    <div className="mb-3">
                                                        <label className="form-label">Occupation</label>
                                                        <input
                                                            name={`references[${index}].occupation`}
                                                            value={ref.occupation}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            className="form-control"
                                                        />
                                                        {typeof touchedReference === "object" &&
                                                            touchedReference?.occupation &&
                                                            typeof errorReference === "object" &&
                                                            errorReference?.occupation && (
                                                                <p className="text-danger small">
                                                                    {errorReference.occupation}
                                                                </p>
                                                            )}
                                                    </div>
                                                    {/* Relationship */}
                                                    <div className="mb-3">
                                                        <label className="form-label">Relationship</label>
                                                        <input
                                                            name={`references[${index}].relationship`}
                                                            value={ref.relationship}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            className="form-control"
                                                        />
                                                        {typeof touchedReference === "object" &&
                                                            touchedReference?.relationship &&
                                                            typeof errorReference === "object" &&
                                                            errorReference?.relationship && (
                                                                <p className="text-danger small">
                                                                    {errorReference.relationship}
                                                                </p>
                                                            )}
                                                    </div>
                                                    {/* Contact */}
                                                    <div>
                                                        <label className="form-label">Contact No & Address</label>
                                                        <textarea
                                                            rows={3}
                                                            name={`references[${index}].contact`}
                                                            value={ref.contact}
                                                            onChange={joiningFormValidation.handleChange}
                                                            onBlur={joiningFormValidation.handleBlur}
                                                            className="form-control"
                                                        />
                                                        {typeof touchedReference === "object" &&
                                                            touchedReference?.contact &&
                                                            typeof errorReference === "object" &&
                                                            errorReference?.contact && (
                                                                <p className="text-danger small">
                                                                    {errorReference.contact}
                                                                </p>
                                                            )}
                                                    </div>
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                                <div data-pdf-keep-together="true">
                                    <h5
                                        className="text-white p-2 rounded mt-4"
                                        style={{ backgroundColor: "#f18200" }}
                                    >
                                        Joining Letter
                                    </h5>
                                    <div className="border rounded p-4 shadow-sm">
                                        {/* Date */}
                                        <p>
                                            Date{" "}
                                            <span className="d-inline-block ms-2" style={{ width: "220px" }}>
                                                <DatePickerInput
                                                    name="joiningLetterDate"
                                                    value={joiningFormValidation.values.joiningLetterDate}
                                                    onValueChange={createDateValueChangeHandler(joiningFormValidation.setFieldValue, "joiningLetterDate")}
                                                    onBlur={joiningFormValidation.handleBlur}
                                                />
                                            </span>
                                            {joiningFormValidation.touched.joiningLetterDate &&
                                                joiningFormValidation.errors.joiningLetterDate && (
                                                    <p className="text-danger small">
                                                        {joiningFormValidation.errors.joiningLetterDate}
                                                    </p>
                                                )}
                                        </p>
                                        <p className="mb-1">To</p>
                                        <p className="mb-0">NAT IT Services Pvt Ltd</p>
                                        <p className="mb-0">Plot No:21, Sruthi Sadan,</p>
                                        <p className="mb-3">Gachibowli, Hyderabad-32</p>
                                        {/* Subject */}
                                        <p>
                                            <strong>Subject:</strong>{" "}
                                            Joining Letter
                                        </p>
                                        <p>Dear Sir/Madam,</p>
                                        <p style={{ lineHeight: 2 }}>
                                            I am pleased to accept your offer and I have honor to inform you that I am
                                            joining in NAT IT Services Pvt Ltd from today (Date)
                                            <span className="d-inline-block mx-2" style={{ width: "220px" }}>
                                                <DatePickerInput
                                                    name="joiningDateText"
                                                    value={joiningFormValidation.values.joiningDateText}
                                                    onValueChange={createDateValueChangeHandler(joiningFormValidation.setFieldValue, "joiningDateText")}
                                                    onBlur={joiningFormValidation.handleBlur}
                                                />
                                            </span>
                                            as a/an
                                            <input
                                                name="designationText"
                                                value={joiningFormValidation.values.designationText}
                                                onChange={joiningFormValidation.handleChange}
                                                className="d-inline-block mx-2 form-control"
                                                style={{ width: "250px" }}
                                            />
                                            . I would be kind enough if you accept this joining letter.
                                        </p>
                                        <p className="mt-4 mb-5">Yours Sincerely,</p>
                                        {/* Signature */}
                                        <p>
                                            Signature{" "}
                                            <span
                                                className="d-inline-block ms-2"
                                                style={{ width: "260px", verticalAlign: 'middle' }}
                                            >
                                                <SignatureUpload
                                                    name="signatureName"
                                                    value={joiningFormValidation.values.signatureName}
                                                    onChange={(value) => {

                                                        joiningFormValidation
                                                            .setFieldValue('signatureName', value)
                                                            .catch(() => undefined);

                                                        onEmployeeSignatureChange?.(
                                                            typeof value === "string"
                                                                ? value
                                                                : URL.createObjectURL(value)
                                                        );

                                                    }}
                                                    onBlur={joiningFormValidation.handleBlur}
                                                />
                                            </span>
                                        </p>
                                        {/* Errors */}
                                        {joiningFormValidation.touched.signatureName &&
                                            joiningFormValidation.errors.signatureName && (
                                                <p className="text-danger small">
                                                    {joiningFormValidation.errors.signatureName}
                                                </p>
                                            )}
                                    </div>
                                </div>
                                {readOnly && (
                                    <div className="alert alert-success mt-4 mb-0">
                                        Your forms have been finally submitted and can no longer be edited.
                                    </div>
                                )}
                                <div className="text-center mt-4">
                                    {!readOnly && (
                                        <Button
                                            type="submit"
                                            className="border-0"
                                            style={{ backgroundColor: "#f18200" }}
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
            </div >
        </>
    );
};
export default JoiningFormalities;
