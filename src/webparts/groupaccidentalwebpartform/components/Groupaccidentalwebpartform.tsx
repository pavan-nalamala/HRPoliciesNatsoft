import * as React from 'react';
import styles from './Groupaccidentalwebpartform.module.scss';
import type { IGroupaccidentalwebpartformProps } from './IGroupaccidentalwebpartformProps';

import { SPHttpClient } from '@microsoft/sp-http';

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
  currentUser: any;
  employeePFData: any[];
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
    completedForms: {} as any,
    sharedEmployeeSignature: '',
    currentUser: null,
    employeePFData: []
  };

  private _isAdminUser(): boolean {
    const email = this.state.employeePFData?.[0]?.EmailID;

    if (!email) return false;

    return email.toLowerCase() === 'hr@natit.in';
  }

  private _getVisibleTabs(): IFormNavItem[] {
    const isGauge = this._isAdminUser();

    const gaugeTabs: FormKey[] = [
      'joiningFormalities',
      'teamLifeInsuranceNomination',
      'gratuityNominationForm',
      'insuranceNominationForm',
      'pfDeclaration'
    ];

    if (isGauge) {
      return FORM_NAV_ITEMS.filter(item =>
        gaugeTabs.indexOf(item.key) !== -1
      );
    }

    return FORM_NAV_ITEMS;
  }

  public componentDidMount(): void {

    if (!document.getElementById('bootstrap-css')) {
      const link = document.createElement('link');
      link.id = 'bootstrap-css';
      link.rel = 'stylesheet';
      link.href =
        'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css';
      document.head.appendChild(link);
    }

    void this.props.context.msGraphClientFactory
      .getClient('3')
      .then(client => client.api('/me').get())
      .then((user: any) => {

        this.setState({ currentUser: user });

        void this._loadEmployeePF(user.mail || user.userPrincipalName);
      })
      .catch(error => console.error('Graph Error:', error));
  }

  private _loadEmployeePF = async (email: string): Promise<void> => {

    try {

      const url =
        "https://natitin.sharepoint.com/sites/NatIt_HRRecruitment" +
        "/_api/web/lists/getbytitle('Candidate Information')/items";

      const response = await this.props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      const data = await response.json();

      const filtered = data.value.filter((item: any) =>
        item.EmailID === email || item.Title === email
      );

      this.setState(
        { employeePFData: filtered },
        () => {
          const isGauge = this._isAdminUser();

          this.setState({
            selectedForm: isGauge
              ? 'joiningFormalities'
              : 'natItServicesHandbook'
          });
        }
      );

    } catch (error) {
      console.error("SharePoint List Error:", error);
    }
  };

  private _handleTabClick = (
    e: React.MouseEvent<HTMLButtonElement>,
    key: FormKey
  ): void => {

    e.preventDefault();
    this.setState({ selectedForm: key });
  };

  private _isFormEnabled(key: FormKey): boolean {

    const index = FORM_NAV_ITEMS.findIndex(i => i.key === key);

    if (index <= 0) return true;

    const prevKey = FORM_NAV_ITEMS[index - 1].key;

    return !!this.state.completedForms[prevKey];
  }

  private _handleFormComplete = (key: FormKey): void => {

    this.setState(prevState => {

      const index = FORM_NAV_ITEMS.findIndex(i => i.key === key);
      const nextKey = FORM_NAV_ITEMS[index + 1]?.key;

      return {
        completedForms: {
          ...prevState.completedForms,
          [key]: true
        },
        selectedForm: nextKey ?? key
      };
    });
  };

  private _handleEmployeeSignatureChange = (value: string): void => {

    this.setState({
      sharedEmployeeSignature: value
    });
  };

  private _renderFormByKey(formKey: FormKey): React.ReactElement {

    const { sharedEmployeeSignature, employeePFData } = this.state;

    const spHttpClient = this.props.context.spHttpClient;
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;

    switch (formKey) {

      case 'joiningFormalities':
        return (
          <JoiningFormalities
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() => this._handleFormComplete('joiningFormalities')}
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'teamLifeInsuranceNomination':
        return (
          <TeamLifeInsuranceNomination
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('teamLifeInsuranceNomination')
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'gratuityNominationForm':
        return (
          <GratuityNominationForm
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('gratuityNominationForm')
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'insuranceNominationForm':
        return (
          <InsuranceNominationForm
            context={this.props.context}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('insuranceNominationForm')
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
          />
        );

      case 'pfDeclaration':
        return (
          <PFDeclaration
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('pfDeclaration')
            }
          />
        );

      case 'natItServicesHrPolicyManual':
        return (
          <NatItServicesHrPolicyManual
            context={this.props.context}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('natItServicesHrPolicyManual')
            }
          />
        );

      case 'natItServicesHandbook':
        return (
          <NatItServicesHandbook
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('natItServicesHandbook')
            }
          />
        );

      default:
        return <div>No Form Found</div>;
    }
  }

  public render(): React.ReactElement<IGroupaccidentalwebpartformProps> {

    const { selectedForm } = this.state;

    return (
      <section>

        <nav className={styles.topNav}>
          {this._getVisibleTabs().map(item => {

            const isActive = selectedForm === item.key;
            const isEnabled = this._isFormEnabled(item.key);

            return (
              <button
                key={item.key}
                type="button"
                className={`${styles.navItem} ${isActive ? styles.navItemActive : ''}`}
                disabled={this._isAdminUser() ? false : !isEnabled}
                onClick={(e) => this._handleTabClick(e, item.key)}
              >
                {item.label}
              </button>
            );
          })}
        </nav>

        {FORM_NAV_ITEMS.map(item => (
          <div
            key={item.key}
            style={{ display: selectedForm === item.key ? 'block' : 'none' }}
          >
            {this._renderFormByKey(item.key)}
          </div>
        ))}

      </section>
    );
  }
}