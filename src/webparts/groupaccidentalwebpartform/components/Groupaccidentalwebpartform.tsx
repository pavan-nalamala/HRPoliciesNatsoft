import * as React from 'react';
import styles from './Groupaccidentalwebpartform.module.scss';
import type { IGroupaccidentalwebpartformProps } from './IGroupaccidentalwebpartformProps';

import {
  GratuityNominationForm,
  InsuranceNominationForm,
  JoiningFormalities,
  NatItServicesHandbook,
  NatItServicesHrPolicyManual,
  PFDeclaration,
  TeamLifeInsuranceNomination
} from './forms';

type FormKey =
  | 'joiningFormalities'
  | 'teamLifeInsuranceNomination'
  | 'gratuityNominationForm'
  | 'insuranceNominationForm'
  | 'pfDeclaration'
  | 'natItServicesHrPolicyManual'
  | 'natItServicesHandbook';

interface IFormNavItem {
  key: FormKey;
  label: string;
}

interface IGroupaccidentalwebpartformState {
  selectedForm: FormKey;
  completedForms: Partial<Record<FormKey, boolean>>;
  sharedEmployeeSignature: string;
}

const FORM_NAV_ITEMS: IFormNavItem[] = [
  { key: 'natItServicesHandbook', label: 'Employee Handbook' },
  { key: 'natItServicesHrPolicyManual', label: 'HR Policy Manual' },
  { key: 'joiningFormalities', label: 'Joining Formalities' },
  { key: 'teamLifeInsuranceNomination', label: 'Team Life Insurance Nomination' },
  { key: 'gratuityNominationForm', label: 'Gratuity Nomination Form' },
  { key: 'insuranceNominationForm', label: 'Insurance Nomination Form' },
  { key: 'pfDeclaration', label: 'PF Declaration' }
];

export default class Groupaccidentalwebpartform extends React.Component<
  IGroupaccidentalwebpartformProps,
  IGroupaccidentalwebpartformState
> {
  public state: IGroupaccidentalwebpartformState = {
    selectedForm: 'natItServicesHandbook',
    completedForms: {},
    sharedEmployeeSignature: ''
  };

  public componentDidMount(): void {
    if (!document.getElementById('bootstrap-css')) {
      const link = document.createElement('link');
      link.id = 'bootstrap-css';
      link.rel = 'stylesheet';
      link.href =
        'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css';
      document.head.appendChild(link);
    }
  }

  private _handleTabClick = (
    e: React.MouseEvent<HTMLButtonElement>,
    key: FormKey
  ): void => {
    e.preventDefault();
    e.stopPropagation();

    if (this.state.selectedForm !== key) {
      this.setState({
        selectedForm: key
      });
    }
  };

  private _isFormEnabled(key: FormKey): boolean {
    const formIndex = FORM_NAV_ITEMS.findIndex((item) => item.key === key);

    if (formIndex <= 0) {
      return true;
    }

    const previousFormKey = FORM_NAV_ITEMS[formIndex - 1].key;
    return !!this.state.completedForms[previousFormKey];
  }

  private _handleFormComplete = (key: FormKey): void => {
    this.setState((prevState) => {
      const currentIndex = FORM_NAV_ITEMS.findIndex((item) => item.key === key);
      const nextFormKey = FORM_NAV_ITEMS[currentIndex + 1]?.key;

      return {
        completedForms: {
          ...prevState.completedForms,
          [key]: true
        },
        selectedForm: nextFormKey ?? key
      };
    });
  };

  private _handleEmployeeSignatureChange = (value: string): void => {
    this.setState({
      sharedEmployeeSignature: value
    });
  };

  private _renderFormByKey(formKey: FormKey): React.ReactElement {
    const { sharedEmployeeSignature } = this.state;

    switch (formKey) {
      case 'joiningFormalities':
        return (
          <JoiningFormalities
            onComplete={() => this._handleFormComplete('joiningFormalities')}
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'teamLifeInsuranceNomination':
        return (
          <TeamLifeInsuranceNomination
            onComplete={() => this._handleFormComplete('teamLifeInsuranceNomination')}
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'gratuityNominationForm':
        return (
          <GratuityNominationForm
            onComplete={() => this._handleFormComplete('gratuityNominationForm')}
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'insuranceNominationForm':
        return (
          <InsuranceNominationForm
            onComplete={() => this._handleFormComplete('insuranceNominationForm')}
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'pfDeclaration':
        return <PFDeclaration onComplete={() => this._handleFormComplete('pfDeclaration')} />;

      case 'natItServicesHrPolicyManual':
        return <NatItServicesHrPolicyManual onComplete={() => this._handleFormComplete('natItServicesHrPolicyManual')} />;

      case 'natItServicesHandbook':
        return <NatItServicesHandbook onComplete={() => this._handleFormComplete('natItServicesHandbook')} />;

      default:
        return (
          <JoiningFormalities
            onComplete={() => this._handleFormComplete('joiningFormalities')}
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );
    }
  }
  
  public render(): React.ReactElement<IGroupaccidentalwebpartformProps> {
    const { selectedForm } = this.state;

    return (
      <section>
        <div>
          <h3 className={styles.pageTitle}>
            HR Forms Dashboard
          </h3>

          {/* Navigation */}
          <nav
            className={styles.topNav}
            aria-label="HR forms navigation"
            style={{
              position: 'relative',
              zIndex: 9999
            }}
          >
            {FORM_NAV_ITEMS.map((item) => {
              const isActive = selectedForm === item.key;
              const isEnabled = this._isFormEnabled(item.key);

              return (
                <button
                  key={item.key}
                  type="button"
                  className={`${styles.navItem} ${
                    isActive ? styles.navItemActive : ''
                  }`}
                  style={{
                    position: 'relative',
                    zIndex: 10000,
                    opacity: isEnabled ? 1 : 0.55,
                    cursor: isEnabled ? 'pointer' : 'not-allowed'
                  }}
                  disabled={!isEnabled}
                  aria-disabled={!isEnabled}
                  onMouseDown={(e) =>
                    this._handleTabClick(e, item.key)
                  }
                >
                  {item.label}
                </button>
              );
            })}
          </nav>

          {/* Form Preview */}
          {FORM_NAV_ITEMS.map((item) => {
            const isVisible = selectedForm === item.key;

            return (
              <div
                key={item.key}
                className={styles.formsPreview}
                style={{
                  position: 'relative',
                  zIndex: 1,
                  display: isVisible ? 'block' : 'none'
                }}
              >
                {this._renderFormByKey(item.key)}
              </div>
            );
          })}
        </div>
      </section>
    );
  }
}
