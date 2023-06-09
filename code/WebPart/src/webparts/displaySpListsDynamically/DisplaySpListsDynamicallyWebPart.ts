import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DisplaySpListsDynamicallyWebPart.module.scss';
import * as strings from 'DisplaySpListsDynamicallyWebPartStrings';

/*Using the HTTP helpers within SharePoint Framework to access the SP Lists.*/
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
/*End of the snippet*/

export interface IDisplaySpListsDynamicallyWebPartProps {
  description: string;
}

/*Retrieving all the SharePoint Lists in the current SharePoint site.*/
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}
/*End of the snippet*/

export default class DisplaySpListsDynamicallyWebPart extends BaseClientSideWebPart<IDisplaySpListsDynamicallyWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {

     /*Replaced the default snippet on 27-05-2023*/
    this.domElement.innerHTML = `
    <section class="${styles.displaySpListsDynamically} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
    <div class="${styles.welcome}">
      <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
      <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
      <div>${this._environmentMessage}</div>
    </div>
    <div>
      <h3>Welcome to SharePoint Framework!</h3>
      <div>Web part description: <strong>${escape(this.properties.description)}</strong></div>
      
      <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
    </div>
    <div id="spListContainer" />
  </section>`;

  this._renderListAsync();
  }
   /*End of snippet*/

   /*This method() retrieve SharePoint Lists from the SharePoint Site inside the WebPart.*/
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      // eslint-disable-next-line @typescript-eslint/no-empty-function
      .catch(() => {});
  }
  /*End of snippet*/

   /*The below method will render all SharePoint Lists in the array ISPList[] into individual 
  items as bullet points outputting with the list title. Basically this will be in a loop 
  so that it will capture all the lists and their behaviours like adding new lists, deleting the exisiting lists.*/
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
  
    /*Below will dynamically add the SharePoint Lists into the Inner HTML of ther webpart.*/
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
     /*End of the snippet*/
  }
   /*End of the snippet*/

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      })
      // eslint-disable-next-line @typescript-eslint/no-empty-function
      .catch(() => {});
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
