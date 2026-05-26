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

const HANDBOOK_FOLDER_URL =
  `${RECRUITMENT_SITE_URL}/Shared%20Documents/JoiningFormalitiesDocuments/NAT%20IT%20SERVICES_Hand%20book.pdf`;

const NatItServicesHandbook = ({
  onComplete,
  employeePFData,
  hasSavedProgress,
  isFinallySubmitted,
  submitButtonLabel = 'Continue'
}: ISequentialFormProps): JSX.Element => {
  const [agreed, setAgreed] = useState(false);
  const [loading, setLoading] = useState(false);
  const readOnly = !!isFinallySubmitted;
  const isAcknowledged = agreed || !!hasSavedProgress;

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
        `${siteUrl}/_api/web/lists/getbytitle('EmployeeHandBook')/items` +
          `?$filter=can_id eq '${canId}'&$top=1&$select=Id`,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      const existingData = await existingRes.json();
      const existingId = existingData.value?.[0]?.Id;

      const body = JSON.stringify({
        Title: 'Employee Handbook Acknowledgement',
        employee_name: employeePFData?.Title || 'Unknown',
        acknowledgement_flag: true,
        acknowledgement_time: new Date().toISOString(),
        can_id: canId
      });

      const url = existingId
        ? `${siteUrl}/_api/web/lists/getbytitle('EmployeeHandBook')/items(${existingId})`
        : `${siteUrl}/_api/web/lists/getbytitle('EmployeeHandBook')/items`;

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
        const err = await response.text();
        throw new Error(err);
      }

      onComplete?.();
    } catch (error) {
      console.error('Submit failed:', error);
      alert('Submission failed');
    }
    setLoading(false);
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
      <FormSection title="NAT IT Services_Handbook" />
      <div className="p-3 bg-light">
        <Card className="shadow border-0 rounded-4 mx-auto">
          <Card.Header
            className="text-white fw-bold fs-4 text-center py-3 border-0"
            style={{ backgroundColor: '#f18200' }}
          >
            NAT IT Services - Employee Handbook
          </Card.Header>
          <Card.Body className="p-4">
            <Alert variant="light" className="border rounded-3 mb-4">
              Please review the Employee Handbook carefully before continuing.
            </Alert>
            <div className="border rounded-4 overflow-hidden shadow-sm mb-4">
              <iframe
                title="Employee Handbook"
                src={HANDBOOK_FOLDER_URL}
                style={{ width: '100%', height: '500px', border: 'none' }}
              />
            </div>
          </Card.Body>
        </Card>
      </div>
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
              label="I have read and understood the Employee Handbook."
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
                style={{
                  backgroundColor: isAcknowledged ? '#f18200' : '#adb5bd',
                  border: 'none',
                  minWidth: '180px'
                }}
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

export default NatItServicesHandbook;
