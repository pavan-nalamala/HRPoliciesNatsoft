import * as React from 'react';

export type ThankYouMessageProps = {
  /** When true, employee is returning after a prior final submit. */
  alreadySubmitted?: boolean;
  statusLabel?: string;
};

const ThankYouMessage = ({
  alreadySubmitted = false,
  statusLabel
}: ThankYouMessageProps): JSX.Element => (
  <div
    className="m-3 text-center"
    role="status"
    aria-live="polite"
    style={{
      minHeight: '360px',
      border: '2px solid #d27ad2',
      borderRadius: '10px',
      background: '#ffffff',
      position: 'relative',
      overflow: 'hidden',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '42px 20px'
    }}
  >
    <div
      style={{
        position: 'absolute',
        top: '-42px',
        left: '-12px',
        width: '72px',
        height: '124px',
        background: '#dca2dd',
        borderRadius: '0 0 24px 24px',
        transform: 'rotate(8deg)'
      }}
    />

    <div style={{ width: '100%', maxWidth: '420px' }}>
      <div
        aria-hidden="true"
        style={{
          height: '180px',
          position: 'relative',
          marginBottom: '18px'
        }}
      >
        <div
          style={{
            position: 'absolute',
            left: '50%',
            bottom: '6px',
            width: '250px',
            height: '34px',
            background: 'rgba(36, 58, 64, 0.07)',
            borderRadius: '50%',
            transform: 'translateX(-50%)'
          }}
        />
        <div
          style={{
            position: 'absolute',
            left: '50%',
            top: '28px',
            transform: 'translateX(-50%) rotate(-8deg)',
            fontSize: '42px',
            lineHeight: 1,
            fontWeight: 800,
            letterSpacing: '0',
            color: '#2e5963',
            textShadow: '6px 0 #1c3f49, 10px 5px rgba(28, 63, 73, 0.26)'
          }}
        >
          THANK
        </div>
        <div
          style={{
            position: 'absolute',
            left: '50%',
            top: '82px',
            transform: 'translateX(-35%) rotate(-8deg)',
            fontSize: '42px',
            lineHeight: 1,
            fontWeight: 800,
            letterSpacing: '0',
            color: '#d15ed1',
            textShadow: '6px 0 #a747b1, 10px 5px rgba(167, 71, 177, 0.22)'
          }}
        >
          YOU
        </div>

        <div style={{ position: 'absolute', left: '58px', top: '82px' }}>
          <div style={{ width: '13px', height: '13px', borderRadius: '50%', background: '#243a40', marginLeft: '15px' }} />
          <div style={{ width: '22px', height: '44px', background: '#d15ed1', borderRadius: '12px 12px 6px 6px', transform: 'rotate(12deg)' }} />
          <div style={{ width: '8px', height: '42px', background: '#274852', borderRadius: '8px', transform: 'rotate(15deg)', marginLeft: '6px', marginTop: '-3px' }} />
          <div style={{ width: '8px', height: '38px', background: '#274852', borderRadius: '8px', transform: 'rotate(-18deg)', marginLeft: '26px', marginTop: '-39px' }} />
        </div>

        <div style={{ position: 'absolute', right: '60px', top: '62px' }}>
          <div style={{ width: '13px', height: '13px', borderRadius: '50%', background: '#243a40', marginLeft: '14px' }} />
          <div style={{ width: '22px', height: '52px', background: '#d15ed1', borderRadius: '12px 12px 6px 6px', transform: 'rotate(-6deg)' }} />
          <div style={{ width: '8px', height: '45px', background: '#274852', borderRadius: '8px', transform: 'rotate(-5deg)', marginLeft: '5px', marginTop: '-2px' }} />
          <div style={{ width: '8px', height: '43px', background: '#274852', borderRadius: '8px', transform: 'rotate(10deg)', marginLeft: '25px', marginTop: '-42px' }} />
        </div>
      </div>

      <h4
        style={{
          color: '#222222',
          fontSize: '22px',
          fontWeight: 700,
          marginBottom: '8px'
        }}
      >
        {alreadySubmitted ? 'Already submitted' : 'Thank you'}
      </h4>
      <p
        style={{
          color: '#7d8790',
          fontSize: '14px',
          marginBottom: 0
        }}
      >
        {alreadySubmitted
          ? `You have already submitted your forms${statusLabel ? ` (${statusLabel})` : ''}. Fields are in view-only mode. Contact HR if you need changes.`
          : 'Your forms have been submitted successfully.'}
      </p>
    </div>
  </div>
);

export default ThankYouMessage;
