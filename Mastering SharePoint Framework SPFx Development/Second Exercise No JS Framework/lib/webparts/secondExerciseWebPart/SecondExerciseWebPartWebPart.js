/* This  file defines the main entry point for the web part */
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
/** CSJ Comment */
/** Property declaration step 01 */
import { PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneDropdown, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './SecondExerciseWebPartWebPart.module.scss';
import * as strings from 'SecondExerciseWebPartWebPartStrings';
import { SPHttpClient } from '@microsoft/sp-http';
var SecondExerciseWebPartWebPart = /** @class */ (function (_super) {
    __extends(SecondExerciseWebPartWebPart, _super);
    function SecondExerciseWebPartWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        return _this;
    }
    /** CSJ Comment */
    /** The render() method is used to render the web part inside that DOM element.
       In the web part, the DOM element is set to a DIV. */
    /** Notice how ${ } is used to output the variable's value in the HTML block.
  * An extra HTML div is used to display this.context.pageContext.web.title. */
    SecondExerciseWebPartWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <section class=\"".concat(styles.secondExerciseWebPart, " ").concat(!!this.context.sdks.microsoftTeams ? styles.teams : '', "\">\n        <div class=\"").concat(styles.welcome, "\">\n          <img alt=\"\" src=\"").concat(this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png'), "\" class=\"").concat(styles.welcomeImage, "\" />\n          <h2>Well done, ").concat(escape(this.context.pageContext.user.displayName), "!</h2>\n          <div>").concat(this._environmentMessage, "</div>\n        </div>\n        <div>\n          <h3>Welcome to SharePoint Framework!</h3>\n          <div>Loading from: <strong>").concat(escape(this.context.pageContext.web.title), "</strong></div>\n          <div>Web part description: <strong>").concat(escape(this.properties.description), "</strong></div>\n          <div>Web part notes: <strong>").concat(escape(this.properties.notes), "</strong></div>\n          <div>Web part status: <strong>").concat(escape(this.properties.statusof.toString()), "</strong></div>\n          <div>Web part comments: <strong>").concat(escape(this.properties.comments), "</strong></div>\n          <div>Web part items: <strong>").concat(escape(this.properties.items.toString()), "</strong></div>\n        </div>\n\n        \n        <div id=\"spListContainer\" />\n\n      </section>");
        /** The below statement list down all the SharePoint lists. */
        this._renderListAsync();
    };
    /** CSJ Comment */
    SecondExerciseWebPartWebPart.prototype._getListData = function () {
        return this.context.spHttpClient.get("".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists?$filter=Hidden eq false"), SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        })
            .catch(function (err) { console.log(err); });
    };
    SecondExerciseWebPartWebPart.prototype._renderList = function (items) {
        var html = '';
        items.forEach(function (item) {
            html += "\n    <ul class=\"".concat(styles.list, "\">\n      <li class=\"").concat(styles.listItem, "\">\n        <span class=\"ms-font-l\">").concat(item.Title, "</span>\n      </li>\n    </ul>");
        });
        if (this.domElement.querySelector('#spListContainer') !== null) {
            this.domElement.querySelector('#spListContainer').innerHTML = html;
        }
    };
    SecondExerciseWebPartWebPart.prototype._renderListAsync = function () {
        var _this = this;
        this._getListData()
            .then(function (response) {
            _this._renderList(response.value);
        })
            .catch(function (err) { console.log(err); });
    };
    SecondExerciseWebPartWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
        });
    };
    SecondExerciseWebPartWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(function (context) {
                var environmentMessage = '';
                switch (context.app.host.name) {
                    case 'Office': // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
                        break;
                    case 'Outlook': // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
                        break;
                    case 'Teams': // running in Teams
                    case 'TeamsModern':
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = strings.UnknownEnvironment;
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
    };
    SecondExerciseWebPartWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    Object.defineProperty(SecondExerciseWebPartWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    /** CSJ Comment */
    /** The property pane is defined in the SecondExerciseNoJsFrameworkWebPart class.
        The getPropertyPaneConfiguration property is where you need to define the property pane. */
    /** When the properties are defined, you can access them in your web part
        by using this.properties.<property-value>, as shown in the render() method: */
    SecondExerciseWebPartWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                    ]
                                }),
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
    };
    return SecondExerciseWebPartWebPart;
}(BaseClientSideWebPart));
export default SecondExerciseWebPartWebPart;
//# sourceMappingURL=SecondExerciseWebPartWebPart.js.map