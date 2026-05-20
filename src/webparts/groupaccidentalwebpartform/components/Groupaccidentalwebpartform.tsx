import * as React from 'react';
import styles from './Groupaccidentalwebpartform.module.scss';
import type { IGroupaccidentalwebpartformProps } from './IGroupaccidentalwebpartformProps';

import { Modal } from 'react-bootstrap';
import { SPHttpClient } from '@microsoft/sp-http';

import {
  GratuityNominationForm,
  InsuranceNominationForm,
  JoiningFormalities,
  NatItServicesHandbook,
  NatItServicesHrPolicyManual,
  PFDeclaration,
  TeamLifeInsuranceNomination,
  ThankYouMessage
} from './forms';

type FormKey =
  | 'joiningFormalities'
  | 'teamLifeInsuranceNomination'
  | 'gratuityNominationForm'
  | 'insuranceNominationForm'
  | 'pfDeclaration'
  | 'natItServicesHrPolicyManual'
  | 'natItServicesHandbook';

type ViewKey = FormKey | 'thankYou';

interface IFormNavItem {
  key: FormKey;
  label: string;
}

interface IFormSubmissionCheck {
  listTitle: string;
  fieldName: string;
  valueType: 'candidateId' | 'email';
}

interface IGroupaccidentalwebpartformState {
  selectedForm: ViewKey;
  completedForms: Partial<Record<FormKey, boolean>>;
  sharedEmployeeSignature: string;
  sharedDateOfBirth: string;
  currentUser: any;
  employeePFData: any[];
  showThankYouModal: boolean;
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

const FORM_SUBMISSION_CHECKS: Record<FormKey, IFormSubmissionCheck> = {
  natItServicesHandbook: {
    listTitle: 'EmployeeHandBook',
    fieldName: 'can_id',
    valueType: 'candidateId'
  },
  natItServicesHrPolicyManual: {
    listTitle: 'HRPolicyManual',
    fieldName: 'can_id',
    valueType: 'candidateId'
  },
  joiningFormalities: {
    listTitle: 'JoiningFormalities',
    fieldName: 'can_id',
    valueType: 'candidateId'
  },
  teamLifeInsuranceNomination: {
    listTitle: 'InsuranceNomination',
    fieldName: 'can_id',
    valueType: 'candidateId'
  },
  gratuityNominationForm: {
    listTitle: 'GratuityNomination',
    fieldName: 'can_id',
    valueType: 'candidateId'
  },
  insuranceNominationForm: {
    listTitle: 'PersonalAccidentalScheme',
    fieldName: 'can_id',
    valueType: 'candidateId'
  },
  pfDeclaration: {
    listTitle: 'EPFDeclarationForm11',
    fieldName: 'SubmittedEmail',
    valueType: 'email'
  }
};

export default class Groupaccidentalwebpartform extends React.Component<
  IGroupaccidentalwebpartformProps,
  IGroupaccidentalwebpartformState
> {

  public state: IGroupaccidentalwebpartformState = {
    selectedForm: 'natItServicesHandbook',
    completedForms: {} as any,
    sharedEmployeeSignature: '',
    sharedDateOfBirth: '',
    currentUser: null,
    employeePFData: [],
    showThankYouModal: false
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

  private _escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private _getNextIncompleteForm(
    completedForms: Partial<Record<FormKey, boolean>>
  ): ViewKey {

    const visibleTabs = this._getVisibleTabs();
    const nextIncomplete = visibleTabs.find(item => !completedForms[item.key]);

    return nextIncomplete?.key ?? 'thankYou';
  }

  private _areAllFormsCompleted(
    completedForms: Partial<Record<FormKey, boolean>>
  ): boolean {

    return FORM_NAV_ITEMS.every(item => !!completedForms[item.key]);
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

          if (isGauge) {
            this.setState({ selectedForm: 'joiningFormalities' });
            return;
          }

          void this._loadCompletedForms(email);
        }
      );

    } catch (error) {
      console.error("SharePoint List Error:", error);
    }
  };

  private _loadCompletedForms = async (email: string): Promise<void> => {

    const candidateId = this.state.employeePFData?.[0]?.ID;

    if (!candidateId) {
      this.setState({ selectedForm: 'natItServicesHandbook' });
      return;
    }

    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    const completedForms: Partial<Record<FormKey, boolean>> = {};

    await Promise.all(
      FORM_NAV_ITEMS.map(async item => {
        const check = FORM_SUBMISSION_CHECKS[item.key];
        const rawValue =
          check.valueType === 'candidateId'
            ? String(candidateId)
            : email;
        const filter =
          `${check.fieldName} eq '${this._escapeODataString(rawValue)}'`;
        const url =
          `${siteUrl}/_api/web/lists/getbytitle('${check.listTitle}')/items` +
          `?$select=Id&$top=1&$filter=${encodeURIComponent(filter)}`;

        try {
          const response = await this.props.context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1
          );

          if (!response.ok) {
            console.error(`Unable to check ${check.listTitle}`);
            return;
          }

          const data = await response.json();
          completedForms[item.key] =
            Array.isArray(data.value) && data.value.length > 0;

        } catch (error) {
          console.error(`Submission check failed for ${check.listTitle}:`, error);
        }
      })
    );

    const nextForm = this._getNextIncompleteForm(completedForms);

    this.setState({
      completedForms,
      selectedForm: nextForm === 'thankYou' ? 'pfDeclaration' : nextForm,
      showThankYouModal: false
    });
  };

  private _handleTabClick = (
    e: React.MouseEvent<HTMLButtonElement>,
    key: FormKey
  ): void => {

    e.preventDefault();
    this.setState({ selectedForm: key });
  };

  private _isFormEnabled(): boolean {

    return true;
  }

  private _handleFormComplete = (key: FormKey): void => {

    this.setState(prevState => {

      const completedForms: Partial<Record<FormKey, boolean>> = {
        ...prevState.completedForms,
        [key]: true
      };

      const nextForm = this._getNextIncompleteForm(completedForms);
      const shouldShowThankYouModal =
        key === 'pfDeclaration' && this._areAllFormsCompleted(completedForms);

      return {
        completedForms,
        selectedForm: nextForm === 'thankYou' ? key : nextForm,
        sharedEmployeeSignature: prevState.sharedEmployeeSignature,
        sharedDateOfBirth: prevState.sharedDateOfBirth,
        showThankYouModal: shouldShowThankYouModal
      };
    });
  };

  private _handleThankYouClose = (): void => {

    this.setState({ showThankYouModal: false });
  };

  private _handleEmployeeSignatureChange = (value: string): void => {

    this.setState({
      sharedEmployeeSignature: value
    });
  };

  private _handleDateOfBirthChange = (value: string): void => {

    this.setState({
      sharedDateOfBirth: value
    });
  };

  private _renderFormByKey(formKey: FormKey): React.ReactElement {

    const { sharedEmployeeSignature, sharedDateOfBirth, employeePFData } = this.state;

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
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={this._handleDateOfBirthChange}
          />
        );

      case 'teamLifeInsuranceNomination':
        return (
          <TeamLifeInsuranceNomination
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('teamLifeInsuranceNomination')
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={this._handleDateOfBirthChange}
          />
        );

      case 'gratuityNominationForm':
        return (
          <GratuityNominationForm
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            onComplete={() =>
              this._handleFormComplete('gratuityNominationForm')
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={this._handleEmployeeSignatureChange}
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={this._handleDateOfBirthChange}
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
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={this._handleDateOfBirthChange}
          />
        );

      case 'pfDeclaration':
        return (
          <PFDeclaration
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            isSubmitted={!!this.state.completedForms.pfDeclaration}
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={this._handleDateOfBirthChange}
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
            isSubmitted={!!this.state.completedForms.natItServicesHrPolicyManual}
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
            isSubmitted={!!this.state.completedForms.natItServicesHandbook}
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

    const { selectedForm, showThankYouModal } = this.state;

    return (
      <section>

        <nav className={styles.topNav}>
          {this._getVisibleTabs().map(item => {

            const isActive = selectedForm === item.key;
            const isEnabled = this._isFormEnabled();

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

        <Modal
          show={showThankYouModal}
          onHide={this._handleThankYouClose}
          centered
        >
          <Modal.Body className="text-center p-4">
            <ThankYouMessage />
          </Modal.Body>
        </Modal>

      </section>
    );
  }
}
