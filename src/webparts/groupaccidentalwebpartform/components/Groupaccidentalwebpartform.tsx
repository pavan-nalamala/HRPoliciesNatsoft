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
import {
  getAdminDeepLinkToken,
  isHrDeepLinkMode,
  isHrUserEmail
} from './forms/hrAccessUtils';
import {
  ADMIN_REVIEW_FORM_KEYS,
  areAllEmployeeFormsSaved,
  arePriorPfFormsSaved,
  FormKey,
  getAdminWorkflowStatus,
  getEmployeeWorkflowStatus,
  getSubmitButtonLabel
} from './forms/formWorkflowUtils';

type ViewKey = FormKey | 'thankYou';

interface IFormNavItem {
  key: FormKey;
  label: string;
}

const RECRUITMENT_SITE_URL =
  'https://natitin.sharepoint.com/sites/NatIt_HRRecruitment';

interface IFormSubmissionCheck {
  listTitle: string;
  fieldName: string;
  valueType: 'candidateId' | 'email';
  siteUrl?: string;
}

interface IGroupaccidentalwebpartformState {
  selectedForm: ViewKey;
  completedForms: Partial<Record<FormKey, boolean>>;
  sharedEmployeeSignature: string;
  sharedDateOfBirth: string;
  currentUser: any;
  employeePFData: any[];
  showThankYouModal: boolean;
  accessDenied: boolean;
  isFinallySubmitted: boolean;
  hrStatus: string;
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
    valueType: 'candidateId',
    siteUrl: RECRUITMENT_SITE_URL
  },
  natItServicesHrPolicyManual: {
    listTitle: 'HRPolicyManual',
    fieldName: 'can_id',
    valueType: 'candidateId',
    siteUrl: RECRUITMENT_SITE_URL
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
    fieldName: 'can_id',
    valueType: 'candidateId'
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
    showThankYouModal: false,
    accessDenied: false,
    isFinallySubmitted: false,
    hrStatus: ''
  };

  // =========================================
  // ADMIN / HR ACCESS
  // =========================================

  private _getLoggedInEmail(): string {
    const user = this.state.currentUser;
    return (
      user?.mail ||
      user?.userPrincipalName ||
      ''
    ).trim();
  }

  private _isHrUser(): boolean {
    return isHrUserEmail(this._getLoggedInEmail());
  }

  /** HR user — admin tabs, employee ID fields, PDF download. */
  private _isAdminUser(): boolean {
    return this._isHrUser();
  }

  private _isAdminDeepLinkSession(): boolean {
    return (
      this._isHrUser() &&
      isHrDeepLinkMode() &&
      !!getAdminDeepLinkToken()
    );
  }

  // =========================================
  // VISIBLE TABS
  // =========================================

  private _getVisibleTabs(): IFormNavItem[] {

    if (this._isAdminDeepLinkSession()) {

      return FORM_NAV_ITEMS.filter(item =>
        ADMIN_REVIEW_FORM_KEYS.indexOf(item.key) !== -1
      );
    }

    return FORM_NAV_ITEMS;
  }

  // =========================================
  // TAB ENABLE
  // =========================================

  private _isFormEnabled(formKey: FormKey): boolean {

    if (this._isAdminUser()) {
      return true;
    }

    if (this.state.isFinallySubmitted) {
      return true;
    }

    const visibleTabs = this._getVisibleTabs();

    if (this.state.completedForms[formKey]) {
      return true;
    }

    const firstIncompleteIndex = visibleTabs.findIndex(
      item => !this.state.completedForms[item.key]
    );

    const clickedIndex = visibleTabs.findIndex(
      item => item.key === formKey
    );

    return (
      firstIncompleteIndex === -1 ||
      clickedIndex <= firstIncompleteIndex
    );
  }

  private _getWorkflowStatus(): 'Pending' | 'Completed' {

    if (this._isAdminDeepLinkSession()) {
      return getAdminWorkflowStatus(this.state.hrStatus);
    }

    return getEmployeeWorkflowStatus(
      this.state.completedForms,
      this.state.isFinallySubmitted
    );
  }

  // =========================================
  // ESCAPE ODATA
  // =========================================

  private _escapeODataString(value: string): string {

    return value.replace(/'/g, "''");
  }

  // =========================================
  // NEXT FORM
  // =========================================

  private _getNextIncompleteForm(
    completedForms: Partial<Record<FormKey, boolean>>
  ): ViewKey {

    const visibleTabs = this._getVisibleTabs();

    const nextIncomplete =
      visibleTabs.find(item => !completedForms[item.key]);

    return nextIncomplete?.key ?? 'thankYou';
  }

  // =========================================
  // ALL FORMS COMPLETED
  // =========================================

  private _areAllFormsCompleted(
    completedForms: Partial<Record<FormKey, boolean>>
  ): boolean {
    return areAllEmployeeFormsSaved(completedForms);
  }

  // =========================================
  // COMPONENT DID MOUNT
  // =========================================

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

        const email = (
          user.mail ||
          user.userPrincipalName ||
          ''
        ).trim();

        const adminToken = getAdminDeepLinkToken();
        const hrDeepLink = isHrDeepLinkMode();

        this.setState({ currentUser: user }, () => {

          if (hrDeepLink) {

            if (!isHrUserEmail(email)) {

              this.setState({ accessDenied: true });
              return;
            }

            if (adminToken) {

              void this._loadEmployeeByAdminToken(adminToken);
              return;
            }
          }

          void this._loadEmployeePF(email);
        });
      })
      .catch(error => console.error('Graph Error:', error));
  }

  // =========================================
  // LOAD EMPLOYEE BY HR DEEP LINK TOKEN
  // =========================================

  private _loadEmployeeByAdminToken = async (
    token: string
  ): Promise<void> => {

    try {

      const siteUrl =
        this.props.context.pageContext.web.absoluteUrl;

      const filter =
        `UniqueToken eq '${this._escapeODataString(token)}'`;

      const epfUrl =
        `${siteUrl}/_api/web/lists/getbytitle('EPFDeclarationForm11')/items` +
        `?$select=Id,can_id,HRStatus&$top=1&$filter=${encodeURIComponent(filter)}`;

      const epfResponse =
        await this.props.context.spHttpClient.get(
          epfUrl,
          SPHttpClient.configurations.v1
        );

      const epfData = await epfResponse.json();

      const epfItem = epfData?.value?.[0];
      const epfHrStatus = epfItem?.HRStatus || '';

      const candidateId = Number(epfItem?.can_id);

      if (!candidateId) {

        console.error('No employee found for admin deep link token.');
        return;
      }

      const recruitmentUrl =
        'https://natitin.sharepoint.com/sites/NatIt_HRRecruitment' +
        `/_api/web/lists/getbytitle('Candidate Information')/items(${candidateId})`;

      const candidateResponse =
        await this.props.context.spHttpClient.get(
          recruitmentUrl,
          SPHttpClient.configurations.v1
        );

      const candidate = await candidateResponse.json();

      const employeeEmail =
        candidate.EmailID || candidate.Title || '';

      this.setState(
        {
          employeePFData: [candidate],
          hrStatus: epfHrStatus
        },
        () => {

          this.setState({
            selectedForm: 'joiningFormalities'
          });

          void this._loadCompletedForms(
            employeeEmail,
            candidateId
          );
        }
      );

    } catch (error) {

      console.error('Admin deep link load error:', error);
    }
  };

  // =========================================
  // LOAD EMPLOYEE PF
  // =========================================

  private _loadEmployeePF = async (
    email: string
  ): Promise<void> => {

    try {

      const url =
        "https://natitin.sharepoint.com/sites/NatIt_HRRecruitment" +
        "/_api/web/lists/getbytitle('Candidate Information')/items";

      const response =
        await this.props.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1
        );

      const data = await response.json();

      const filtered = data.value.filter((item: any) =>
        item.EmailID === email ||
        item.Title === email
      );

      this.setState(
        { employeePFData: filtered },
        () => {
          void this._loadCompletedForms(
            email,
            filtered[0]?.ID
          );
        }
      );

    } catch (error) {

      console.error(
        "SharePoint List Error:",
        error
      );
    }
  };

  // =========================================
  // LOAD COMPLETED FORMS
  // =========================================

  private _loadCompletedForms = async (
    email: string,
    candidateIdOverride?: number
  ): Promise<void> => {

    const candidateId =
      candidateIdOverride ??
      this.state.employeePFData?.[0]?.ID;

    if (!candidateId) {

      this.setState({
        selectedForm: 'natItServicesHandbook'
      });

      return;
    }

    const defaultSiteUrl =
      this.props.context.pageContext.web.absoluteUrl;

    const completedForms:
      Partial<Record<FormKey, boolean>> = {};

    let isFinallySubmitted = false;
    let hrStatus = '';

    await Promise.all(
      FORM_NAV_ITEMS.map(async item => {

        const check =
          FORM_SUBMISSION_CHECKS[item.key];

        const rawValue =
          check.valueType === 'candidateId'
            ? String(candidateId)
            : email;

        const filter =
          `${check.fieldName} eq '${this._escapeODataString(rawValue)}'`;

        const selectFields =
          item.key === 'pfDeclaration'
            ? 'Id,HRStatus,SubmittedTime'
            : 'Id';

        const listSiteUrl = check.siteUrl || defaultSiteUrl;

        const url =
          `${listSiteUrl}/_api/web/lists/getbytitle('${check.listTitle}')/items` +
          `?$select=${selectFields}&$top=1&$filter=${encodeURIComponent(filter)}`;

        try {

          const response =
            await this.props.context.spHttpClient.get(
              url,
              SPHttpClient.configurations.v1
            );

          if (!response.ok) {

            console.error(
              `Unable to check ${check.listTitle}`
            );

            return;
          }

          const data =
            await response.json();

          const row = data.value?.[0];

          completedForms[item.key] =
            Array.isArray(data.value) &&
            data.value.length > 0;

          if (item.key === 'pfDeclaration' && row) {

            hrStatus = row.HRStatus || '';

            isFinallySubmitted =
              hrStatus === 'Pending' ||
              hrStatus === 'Completed' ||
              !!row.SubmittedTime;
          }

        } catch (error) {

          console.error(
            `Submission check failed for ${check.listTitle}:`,
            error
          );
        }
      })
    );

    const nextForm =
      this._getNextIncompleteForm(completedForms);

    this.setState({
      completedForms,
      isFinallySubmitted,
      hrStatus,
      selectedForm:
        nextForm === 'thankYou'
          ? 'pfDeclaration'
          : nextForm,
      showThankYouModal: false
    });
  };

  // =========================================
  // TAB CLICK
  // =========================================

  private _handleTabClick = (
    e: React.MouseEvent<HTMLButtonElement>,
    key: FormKey
  ): void => {

    e.preventDefault();

    this.setState({
      selectedForm: key
    });
  };

  // =========================================
  // FORM COMPLETE
  // =========================================

  private _handleFormComplete = (
    key: FormKey
  ): void => {

    this.setState(prevState => {

      const completedForms:
        Partial<Record<FormKey, boolean>> = {

        ...prevState.completedForms,
        [key]: true
      };

      const nextForm =
        this._getNextIncompleteForm(completedForms);

      const isFinalPfSubmit =
        key === 'pfDeclaration' && !this._isHrUser();

      const shouldShowThankYouModal =
        isFinalPfSubmit &&
        this._areAllFormsCompleted(completedForms);

      return {

        completedForms,

        isFinallySubmitted:
          prevState.isFinallySubmitted || isFinalPfSubmit,

        hrStatus:
          key === 'pfDeclaration' && this._isHrUser()
            ? 'Completed'
            : isFinalPfSubmit
              ? 'Pending'
              : prevState.hrStatus,

        selectedForm:
          nextForm === 'thankYou'
            ? key
            : nextForm,

        sharedEmployeeSignature:
          prevState.sharedEmployeeSignature,

        sharedDateOfBirth:
          prevState.sharedDateOfBirth,

        showThankYouModal:
          shouldShowThankYouModal
      };
    });
  };

  // =========================================
  // THANK YOU CLOSE
  // =========================================

  private _handleThankYouClose(): void {

    this.setState({
      showThankYouModal: false
    });
  }

  // =========================================
  // SHARED SIGNATURE
  // =========================================

  private _handleEmployeeSignatureChange(
    value: string
  ): void {

    this.setState({
      sharedEmployeeSignature: value
    });
  }

  // =========================================
  // SHARED DOB
  // =========================================

  private _handleDateOfBirthChange(
    value: string
  ): void {

    this.setState({
      sharedDateOfBirth: value
    });
  }

  // =========================================
  // RENDER FORM
  // =========================================

  private _renderFormByKey(
    formKey: FormKey
  ): React.ReactElement {

    const {
      sharedEmployeeSignature,
      sharedDateOfBirth,
      employeePFData
    } = this.state;

    const spHttpClient =
      this.props.context.spHttpClient;

    const siteUrl =
      this.props.context.pageContext.web.absoluteUrl;

    const candidateId =
      employeePFData?.[0]?.ID;

    const isAdminEdit = this._isHrUser();
    const isFinallySubmitted = this.state.isFinallySubmitted;
    const hasSavedProgress = !!this.state.completedForms[formKey];
    const workflowStatus = this._getWorkflowStatus();
    const submitButtonLabel = getSubmitButtonLabel(formKey);
    const canSubmitFinal =
      formKey === 'pfDeclaration' &&
      arePriorPfFormsSaved(this.state.completedForms);

    const formProps = {
      isAdminEdit,
      isFinallySubmitted,
      hasSavedProgress,
      workflowStatus,
      submitButtonLabel,
      canSubmitFinal
    };

    switch (formKey) {

      case 'joiningFormalities':

        return (
          <JoiningFormalities
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            candidateId={candidateId}
            {...formProps}
            onComplete={() =>
              this._handleFormComplete(
                'joiningFormalities'
              )
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={
              this._handleEmployeeSignatureChange
            }
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={
              this._handleDateOfBirthChange
            }
          />
        );

      case 'teamLifeInsuranceNomination':

        return (
          <TeamLifeInsuranceNomination
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            candidateId={candidateId}
            {...formProps}
            onComplete={() =>
              this._handleFormComplete(
                'teamLifeInsuranceNomination'
              )
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={
              this._handleEmployeeSignatureChange
            }
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={
              this._handleDateOfBirthChange
            }
          />
        );

      case 'gratuityNominationForm':

        return (
          <GratuityNominationForm
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            candidateId={candidateId}
            {...formProps}
            onComplete={() =>
              this._handleFormComplete(
                'gratuityNominationForm'
              )
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={
              this._handleEmployeeSignatureChange
            }
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={
              this._handleDateOfBirthChange
            }
          />
        );

      case 'insuranceNominationForm':

        return (
          <InsuranceNominationForm
            context={this.props.context}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            candidateId={candidateId}
            {...formProps}
            onComplete={() =>
              this._handleFormComplete(
                'insuranceNominationForm'
              )
            }
            sharedEmployeeSignature={sharedEmployeeSignature}
            onEmployeeSignatureChange={
              this._handleEmployeeSignatureChange
            }
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={
              this._handleDateOfBirthChange
            }
          />
        );

      case 'pfDeclaration':

        return (
          <PFDeclaration
            context={this.props.context}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            candidateId={candidateId}
            {...formProps}
            sharedDateOfBirth={sharedDateOfBirth}
            onSharedDateOfBirthChange={
              this._handleDateOfBirthChange
            }
            onComplete={() =>
              this._handleFormComplete(
                'pfDeclaration'
              )
            }
          />
        );

      case 'natItServicesHrPolicyManual':

        return (
          <NatItServicesHrPolicyManual
            context={this.props.context}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            {...formProps}
            onComplete={() =>
              this._handleFormComplete(
                'natItServicesHrPolicyManual'
              )
            }
          />
        );

      case 'natItServicesHandbook':

        return (
          <NatItServicesHandbook
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            employeePFData={employeePFData[0]}
            {...formProps}
            onComplete={() =>
              this._handleFormComplete(
                'natItServicesHandbook'
              )
            }
          />
        );

      default:

        return <div>No Form Found</div>;
    }
  }

  // =========================================
  // RENDER
  // =========================================

  public render():
    React.ReactElement<IGroupaccidentalwebpartformProps> {

    const {
      selectedForm,
      showThankYouModal,
      accessDenied
    } = this.state;

    if (accessDenied) {

      return (
        <section className="p-4">
          <div className="alert alert-danger mb-0" role="alert">
            This link is restricted to HR administrators. Please sign in with an
            authorized HR account to review employee submissions.
          </div>
        </section>
      );
    }

    return (
      <section>

        {this._isAdminDeepLinkSession() && (
          <div className="alert alert-info mx-3 mt-3 mb-0" role="status">
            HR review mode — verify employee data, complete Employee ID fields,
            and download PDFs as needed.
          </div>
        )}

        {this.state.employeePFData.length > 0 && (
          <div className="mx-3 mt-3 mb-0 d-flex align-items-center gap-2">
            <span className="text-muted">Status:</span>
            <span
              className={`badge ${this._getWorkflowStatus() === 'Completed' ? 'bg-success' : 'bg-warning text-dark'}`}
            >
              {this._getWorkflowStatus()}
            </span>
          </div>
        )}

        <nav className={styles.topNav}>

          {this._getVisibleTabs().map(item => {

            const isActive =
              selectedForm === item.key;

            const isEnabled =
              this._isFormEnabled(item.key);

            return (
              <button
                key={item.key}
                type="button"
                className={`${styles.navItem} ${isActive ? styles.navItemActive : ''}`}
                disabled={
                  this._isAdminUser()
                    ? false
                    : !isEnabled
                }
                onClick={(e) =>
                  this._handleTabClick(
                    e,
                    item.key
                  )
                }
              >
                {item.label}
              </button>
            );
          })}
        </nav>

        {FORM_NAV_ITEMS.map(item => (

          <div
            key={item.key}
            style={{
              display:
                selectedForm === item.key
                  ? 'block'
                  : 'none'
            }}
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