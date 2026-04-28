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

const HANDBOOK_PDF_URL = 'https://natitin.sharepoint.com/sites/NatIt_HRRecruitment/Shared%20Documents/JoiningFormalitiesDocuments/NAT%20IT%20SERVICES_Hand%20book.pdf';

const NatItServicesHandbook = ({ onComplete }: ISequentialFormProps): JSX.Element => {
  const [agreed, setAgreed] = useState(false);

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

          {/* Header */}
          <Card.Header
            className="text-white fw-bold fs-4 text-center py-3 border-0"
            style={{ backgroundColor: '#f18200' }}
          >
            NAT IT Services - Employee Handbook
          </Card.Header>

          <Card.Body className="p-4">

            {/* Info Alert */}
            <Alert
              variant="light"
              className="border rounded-3 mb-4"
            >
              Please review the Employee Handbook carefully before continuing.
              You must accept the handbook acknowledgment.
            </Alert>

            {/* PDF Viewer */}
            <div
              className="border rounded-4 overflow-hidden shadow-sm mb-4"
              style={{ backgroundColor: '#fff' }}
            >
              <iframe
                title="NAT IT Services Handbook"
                src={HANDBOOK_PDF_URL}
                style={{
                  width: '100%',
                  height: '500px',
                  border: 'none'
                }}
              />
            </div>

          </Card.Body>
        </Card>
      </div>

      {/* Fixed Bottom Bar */}
      <div
        style={{
          position: 'fixed',
          bottom: 0,
          left: 0,
          width: '100%',
          zIndex: 1050,
          background: '#ffffff',
          borderTop: '1px solid #dee2e6',
          boxShadow: '0 -4px 12px rgba(0,0,0,0.08)',
          padding: '12px 20px'
        }}
      >
        <div style={{ maxWidth: '1100px', margin: '0 auto' }}>
          <Row className="align-items-center g-3">
            <Col md={8}>
              <Form.Check
                type="checkbox"
                id="agreeCheck"
                label="I have read and understood the NAT IT Services Employee Handbook."
                checked={agreed}
                onChange={(e) => setAgreed(e.target.checked)}
                className="fw-medium"
              />
            </Col>

            <Col md={4} className="text-md-end text-start">
              <Button
                size="lg"
                disabled={!agreed}
                onClick={() => onComplete?.()}
                style={{
                  backgroundColor: agreed ? '#f18200' : '#adb5bd',
                  border: 'none',
                  minWidth: '180px'
                }}
              >
                Continue
              </Button>
            </Col>
          </Row>
        </div>
      </div>

      {/* Bottom Space */}
      {/* <div style={{ height: '100px' }}></div> */}
    </>
  );
};

export default NatItServicesHandbook;
