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

const RECRUITMENT_SITE_URL =
    'https://natitin.sharepoint.com/sites/NatIt_HRRecruitment';

const NatItServicesHrPolicyManual = ({
    onComplete,
    context,
    employeePFData,
    hasSavedProgress,
    isFinallySubmitted,
    submitButtonLabel = 'Continue'
}: ISequentialFormProps): JSX.Element => {

    const [agreed, setAgreed] = useState(false);
    const [loading, setLoading] = useState(false);
    const readOnly = !!isFinallySubmitted;
    const isAcknowledged = agreed || !!hasSavedProgress;

    const PDF_URL =
        `${context?.pageContext.web.absoluteUrl}/Shared Documents/JoiningFormalitiesDocuments/Nat IT Services _HR Policy Manual 1.0.pdf`;

    useEffect(() => {
        if (!employeePFData?.ID || !hasSavedProgress) {
            return;
        }
        setAgreed(true);
    }, [employeePFData?.ID, hasSavedProgress]);

    const submitAcknowledgement = async (): Promise<void> => {
        setLoading(true);

        try {
            const siteUrl = RECRUITMENT_SITE_URL;
            const canId = String(employeePFData?.ID);
            const digestRes = await fetch(`${siteUrl}/_api/contextinfo`, {
                method: 'POST',
                headers: { Accept: 'application/json;odata=nometadata' }
            });
            const digestData = await digestRes.json();
            const digest = digestData.FormDigestValue;

            const existingRes = await fetch(
                `${siteUrl}/_api/web/lists/getbytitle('HRPolicyManual')/items` +
                    `?$filter=can_id eq '${canId}'&$top=1&$select=Id`,
                { headers: { Accept: 'application/json;odata=nometadata' } }
            );
            const existingData = await existingRes.json();
            const existingId = existingData.value?.[0]?.Id;

            const body = JSON.stringify({
                Title: 'HR Policy Manual Acknowledgement',
                employee_name: employeePFData?.Title || 'Unknown',
                acknowledgement_flag: true,
                acknowledgement_time: new Date().toISOString(),
                can_id: canId
            });

            const url = existingId
                ? `${siteUrl}/_api/web/lists/getbytitle('HRPolicyManual')/items(${existingId})`
                : `${siteUrl}/_api/web/lists/getbytitle('HRPolicyManual')/items`;

            const response = await fetch(url, {
                method: existingId ? 'MERGE' : 'POST',
                headers: {
                    Accept: 'application/json;odata=nometadata',
                    'Content-Type': 'application/json;odata=nometadata',
                    'X-RequestDigest': digest,
                    ...(existingId ? { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' } : {})
                },
                body
            });

            if (!response.ok) {
                throw new Error(await response.text());
            }

            onComplete?.();
        } catch (error) {
            console.error('Submission Error:', error);
            alert('Submission failed');
        } finally {
            setLoading(false);
        }
    };

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

            <div
                style={{
                    position: 'fixed',
                    bottom: 0,
                    left: 0,
                    width: '100%',
                    zIndex: 1050,
                    background: '#fff',
                    borderTop: '1px solid #dee2e6',
                    padding: '12px 20px'
                }}
            >
                <Row className="align-items-center">
                    <Col md={8}>
                        <Form.Check
                            type="checkbox"
                            label="I have read and agree to the HR Policy Manual"
                            checked={isAcknowledged}
                            disabled={readOnly}
                            onChange={(e) => setAgreed(e.target.checked)}
                        />
                    </Col>

                    <Col md={4} className="text-end">
                        {!readOnly && (
                            <Button
                                disabled={!isAcknowledged || loading}
                                onClick={submitAcknowledgement}
                                style={{ backgroundColor: '#f18200', border: 'none' }}
                            >
                                {loading ? 'Saving...' : submitButtonLabel}
                            </Button>
                        )}
                        {readOnly && (
                            <span className="text-success fw-semibold">Completed</span>
                        )}
                    </Col>
                </Row>
            </div>
        </>
    );
};

export default NatItServicesHrPolicyManual;
