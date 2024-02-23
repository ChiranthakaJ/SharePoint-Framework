/* This  file defines the main entry point for the web part */

import { Version } from '@microsoft/sp-core-library';
import {

  /** CSJ Comment */
  /** Property declaration step 01 */

  type IPropertyPaneConfiguration,
PropertyPaneTextField,
PropertyPaneCheckbox,
PropertyPaneDropdown,
PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'; 

/** CSJ Comment*/
/** BaseClientSideWebPart implements the minimal functionality that is required to build a web part.*/
/** This class also provides many parameters to validate and access read-only properties such as displayMode, 
web part properties, web part context, web part instanceId, the web part domElement*/

import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SecondExerciseNoJsFrameworkWebPart.module.scss';
import * as strings from 'SecondExerciseNoJsFrameworkWebPartStrings';

/** CSJ Comment */
/** The property type is defined as an interface before the class.*/
/** This property definition is used to define custom property types for your web part, 
 * which is described in the property pane section later. */

export interface ISecondExerciseNoJsFrameworkWebPartProps {
  description: string;
  complete_details: string;
  live: boolean;
  available_features: string;
  undergraduate: boolean;
}

export default class SecondExerciseNoJsFrameworkWebPart extends BaseClientSideWebPart<ISecondExerciseNoJsFrameworkWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  /** CSJ Comment */
  /** The render() method is used to render the web part inside that DOM element. 
   In the web part, the DOM element is set to a DIV. */
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.secondExerciseNoJsFramework} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">s
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.name)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
      </div>
    </section>`;
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** CSJ Comment */
  /** The property pane is defined in the SecondExerciseNoJsFrameworkWebPart class. 
      The getPropertyPaneConfiguration property is where you need to define the property pane. */
  /** When the properties are defined, you can access them in your web part 
      by using this.properties.<property-value>, as shown in the render() method: */

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
                label: 'Description'
              }),
              PropertyPaneTextField('complete_details', {
                label: 'Multi-line Text Field',
                multiline: true
              }),
              PropertyPaneCheckbox('live', {
                text: 'Checkbox'
              }),
              PropertyPaneDropdown('available_features', {
                label: 'Dropdown',
                options: [
                  { key: '1', text: 'One' },
                  { key: '2', text: 'Two' },
                  { key: '3', text: 'Three' },
                  { key: '4', text: 'Four' }
                ]}),
              PropertyPaneToggle('graduate', {
                label: 'Toggle',
                onText: 'On',
                offText: 'Off'
              })
            ]
        }
      ]
    },
  }