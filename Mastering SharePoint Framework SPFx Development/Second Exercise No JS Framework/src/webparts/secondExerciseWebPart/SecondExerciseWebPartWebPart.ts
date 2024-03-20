/* This  file defines the main entry point for the web part */

import { Version } from '@microsoft/sp-core-library';

/** CSJ Comment */
  /** Property declaration step 01 */
import {
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

import styles from './SecondExerciseWebPartWebPart.module.scss';
import * as strings from 'SecondExerciseWebPartWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

/** CSJ Comment */
/** The property type is defined as an interface before the class.*/
/** This property definition is used to define custom property types for your web part, 
 * which is described in the property pane section later. */

export interface ISecondExerciseWebPartWebPartProps {
  description: string;
  notes: string;
  statusof: boolean;
  comments: string;
  items: boolean;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class SecondExerciseWebPartWebPart extends BaseClientSideWebPart<ISecondExerciseWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

 /** CSJ Comment */
/** The render() method is used to render the web part inside that DOM element. 
   In the web part, the DOM element is set to a DIV. */
     /** Notice how ${ } is used to output the variable's value in the HTML block. 
   * An extra HTML div is used to display this.context.pageContext.web.title. */
 
  public render(): void {
    this.domElement.innerHTML = `
      <section class="${styles.secondExerciseWebPart} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.welcome}">
          <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <div>${this._environmentMessage}</div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
          <div>Web part description: <strong>${escape(this.properties.description)}</strong></div>
          <div>Web part notes: <strong>${escape(this.properties.notes)}</strong></div>
          <div>Web part status: <strong>${escape(this.properties.statusof.toString())}</strong></div>
          <div>Web part comments: <strong>${escape(this.properties.comments)}</strong></div>
          <div>Web part items: <strong>${escape(this.properties.items.toString())}</strong></div>
        </div>

        
        <div id="spListContainer" />

      </section>`;
      /** The below statement list down all the SharePoint lists. */
      this._renderListAsync();
  }

  /** CSJ Comment */
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .catch((err) => { console.log(err); });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
    <ul class="${styles.list}">
      <li class="${styles.listItem}">
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });
  
    if(this.domElement.querySelector('#spListContainer') !== null) {
      this.domElement.querySelector('#spListContainer')!.innerHTML = html;
    }
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      })
      .catch((err) => { console.log(err); });
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
              /**header: {
                description: strings.PropertyPaneDescription
              },*/
              groups: [
                {
                  groupName: strings.BasicGroupName,
                  groupFields: [
                  PropertyPaneTextField('description', {
                    label: 'Description'
                  }),
                  PropertyPaneTextField('notes', {
                    label: 'Multi-line Text Field',
                    multiline: true
                  }),
                  PropertyPaneCheckbox('comments', {
                    text: 'Checkbox'
                  }),
                  PropertyPaneDropdown('items', {
                    label: 'Dropdown',
                    options: [
                      { key: '1', text: 'One' },
                      { key: '2', text: 'Two' },
                      { key: '3', text: 'Three' },
                      { key: '4', text: 'Four' }
                    ]}),
                  PropertyPaneToggle('statusof', {
                    label: 'Toggle',
                    onText: 'On',
                    offText: 'Off'
                  })
                ]
                }
              ]
            }
          ]
        };
      }
}
