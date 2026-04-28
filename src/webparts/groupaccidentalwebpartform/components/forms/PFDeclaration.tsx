import * as React from "react";
import {
    Container,
    Card,
    Form,
    Row,
    Col,
    Button,
    Alert,
    Modal,
} from "react-bootstrap";
import { useFormik } from "formik";
import * as Yup from "yup";
import type { ISequentialFormProps } from "./ISequentialFormProps";
import {
    createDateValueChangeHandler,
    createDateValidation,
    DatePickerInput,
    formatDateForDisplay
} from "./dateFieldUtils";

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

export default function PFDeclaration({ onComplete }: ISequentialFormProps): JSX.Element {
    const [showSuccess, setShowSuccess] = React.useState(false);
    const [submitted, setSubmitted] = React.useState(false);

    const validationSchema = Yup.object().shape({
        employeeName: Yup.string().required("Employee Name is required"),
        dateOfBirth: createDateValidation("Date of Birth is required", "Date of Birth must be in DD/MM/YYYY format"),
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
        onSubmit: async (values) => {
            console.log("values", values);
            setSubmitted(true);
            setShowSuccess(true);
            onComplete?.();
        },
    });

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

    if (submitted) {
        return (
            <Container className="py-5">
                <Card className="shadow rounded-4 text-center p-5">
                    <h3>You have submitted this form!</h3>
                    <Button href="#" className="mt-3">
                        My Request
                    </Button>
                </Card>
            </Container>
        );
    }

    return (
        <Container fluid className="py-4 bg-light">
            <Card
                className="mx-auto shadow border-0 rounded-4"
                style={{ maxWidth: 1150 }}
            >
                <Card.Header
                    className="text-white text-center fw-bold fs-4"
                    style={{ background: "#f18200" }}
                >
                    Form No.11 Declaration Form
                </Card.Header>
                <Card.Body className="p-4">
                    <Alert variant="light">
                        Employees Provident Fund Organisation Declaration Form
                    </Alert>
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
                                    />
                                    <Form.Check
                                        type="radio"
                                        name="relationType"
                                        label="Spouse"
                                        value="Spouse"
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
                                        <option>Transgender</option>
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
                                        <option>Married</option>
                                        <option>Unmarried</option>
                                        <option>Widow</option>
                                        <option>Widower</option>
                                        <option>Divorcee</option>
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
                                        <option>Illiterate</option>
                                        <option>Non-Matric</option>
                                        <option>Matric</option>
                                        <option>Senior Secondary</option>
                                        <option>Graduate</option>
                                        <option>Post Graduate</option>
                                        <option>Doctor</option>
                                        <option>Technical/Professional</option>
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

                        <h5 className="mt-4">KYC Details</h5>
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
                                    className="mt-3"
                                    name="declarationAccepted"
                                    label="I acknowledge the information provided is correct."
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

                                <Card className="mt-4 ">
                                    <Card.Header
                                        className="fw-bold text-center text-uppercase"
                                        style={{
                                            backgroundColor: "#e9ecef",
                                            fontSize: "18px"
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
                                                    label="The KYC details of the above member in the UAN database have not been uploaded"
                                                    className="mb-2"
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
                                                    label="Have been uploaded but not approved"
                                                    className="mb-2"
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
                                                    label="Have been uploaded and approved with DSC"
                                                    className="mb-2"
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
                                                    label="The KYC details of the above member in the UAN database have been approved with Digital Signature Certificate and transfer request has been generated on portal."
                                                    className="mb-2"
                                                    checked={formik.values.employerTransferApproved}
                                                    onChange={handleExclusiveCheckboxChange("employerTransferApproved", [
                                                        "employerTransferApproved",
                                                        "employerPhysicalClaim",
                                                    ])}
                                                />

                                                <Form.Check
                                                    name="employerPhysicalClaim"
                                                    type="checkbox"
                                                    label="As the DSC of establishment are not registered with EPFO, the member has been informed to file physical claim (Form-13) for transfer of funds from his previous establishment."
                                                    className="mb-2"
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
                            </Card.Body>
                        </Card>
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
                            >
                                Submit
                            </Button>
                        </div>
                    </Form>
                </Card.Body>
            </Card>

            <Modal show={showSuccess} onHide={() => setShowSuccess(false)} centered>
                <Modal.Body className="text-center p-5">
                    <h3>Thank You!</h3>
                    <p>Your details have been successfully submitted.</p>
                    <Button onClick={() => setShowSuccess(false)}>Close</Button>
                </Modal.Body>
            </Modal>
        </Container>
    );
}
