import * as React from 'react';
import { useEffect, useState } from 'react';
import FormSection from './FormSection';
import {
    Card,
    Form,
    Button,
    Row,
    Col,
    Alert
} from 'react-bootstrap';

import type { ISequentialFormProps } from './ISequentialFormProps';
import { getSP } from '../../../../pnpjsConfig';

const NatItServicesHrPolicyManual = ({
    onComplete,
    context,
    employeePFData
}: ISequentialFormProps & { employeePFData: any }): JSX.Element => {

    const [agreed, setAgreed] = useState(false);
    const [loading, setLoading] = useState(false);

    const sp = React.useMemo(() => getSP(context), [context]);

    const PDF_URL =
        `${context?.pageContext.web.absoluteUrl}/Shared Documents/JoiningFormalitiesDocuments/Nat IT Services _HR Policy Manual 1.0.pdf`;

    React.useEffect(() => {

        if (employeePFData) {


            console.log(employeePFData);


        }

    }, [employeePFData]);
    
    // SUBMIT TO SHAREPOINT (PnP)
    
    const submitAcknowledgement = async (): Promise<void> => {
        setLoading(true);

        try {
            await sp.web.lists.getByTitle("HRPolicyManual").items.add({
                Title: "HR Policy Manual Acknowledgement",
                employee_name: employeePFData?.Title || "Unknown",
                can_id: String(employeePFData?.ID),
                acknowledgement_flag: true,
                acknowledgement_time: new Date()
            });

            onComplete?.();

        } catch (error) {
            console.error("Submission Error:", error);
            alert("Submission failed");
        } finally {
            setLoading(false);
        }
    };

    
    // LOAD BOOTSTRAP
    
    useEffect(() => {
        if (!document.getElementById('bootstrap-css')) {
            const link = document.createElement('link');
            link.id = 'bootstrap-css';
            link.rel = 'stylesheet';
            link.href =
                'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css';
            document.head.appendChild(link);
        }
    }, []);

    return (
        <>
            <FormSection title="HR Policy Manual" />

            <Card className="shadow border-0 rounded-4 mx-auto">
                <Card.Header
                    className="text-white text-center fw-bold"
                    style={{ backgroundColor: '#f18200' }}
                >
                    HR Policy Manual
                </Card.Header>

                <Card.Body>
                    <Alert variant="light">
                        Please read and accept the policy before continuing.
                    </Alert>

                    <iframe
                        title="HR Policy Manual"
                        src={PDF_URL}
                        style={{ width: '100%', height: '500px', border: 'none' }}
                    />
                </Card.Body>
            </Card>

            {/* FOOTER ACTION */}

            <div className="p-3"
                style={{
                    position: 'fixed',
                    bottom: 0,
                    left: 0,
                    width: '100%',
                    zIndex: 1050,
                    background: '#fff',
                    borderTop: '1px solid #dee2e6',
                    padding: '12px 20px'
                }}>
                <Row className="align-items-center">
                    <Col md={8}>
                        <Form.Check
                            type="checkbox"
                            label="I have read and agree to the HR Policy Manual"
                            checked={agreed}
                            onChange={(e) => setAgreed(e.target.checked)}
                        />
                    </Col>

                    <Col md={4} className="text-end">
                        <Button
                            disabled={!agreed || loading}
                            onClick={submitAcknowledgement}
                            style={{ backgroundColor: '#f18200', border: 'none' }}
                        >
                            {loading ? "Submitting..." : "Continue"}
                        </Button>
                    </Col>
                </Row>
            </div>
        </>
    );
};

export default NatItServicesHrPolicyManual;