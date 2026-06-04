import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import strings from 'GroupaccidentalwebpartformWebPartStrings';
import Groupaccidentalwebpartform from './components/Groupaccidentalwebpartform';
import { IGroupaccidentalwebpartformProps } from './components/IGroupaccidentalwebpartformProps';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IGroupaccidentalwebpartformWebPartProps {
  description: string;
  hrDocumentLibraryTitle: string;
  hrFormPagePath: string;
}

export default class GroupaccidentalwebpartformWebPart extends BaseClientSideWebPart<IGroupaccidentalwebpartformWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IGroupaccidentalwebpartformProps> = React.createElement(
      Groupaccidentalwebpartform,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        hrDocumentLibraryTitle:
          this.properties.hrDocumentLibraryTitle || 'Link Document Library',
        hrFormPagePath: this.properties.hrFormPagePath
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('hrDocumentLibraryTitle', {
                  label: 'HR document library name',
                  description:
                    'Documents are archived here after HR verification (default: Link Document Library).'
                }),
                PropertyPaneTextField('hrFormPagePath', {
                  label: 'HR form page path',
                  description:
                    'Server-relative path for AdminDeepLink emails (default: /sites/NatIt_HRRecruitment/SitePages/GroupPersonalAccidentalNomination.aspx).'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
