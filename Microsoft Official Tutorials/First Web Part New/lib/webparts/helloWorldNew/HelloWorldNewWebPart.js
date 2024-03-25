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
import { PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneDropdown, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './HelloWorldNewWebPart.module.scss';
import * as strings from 'HelloWorldNewWebPartStrings';
import { SPHttpClient } from '@microsoft/sp-http';
var HelloWorldNewWebPart = /** @class */ (function (_super) {
    __extends(HelloWorldNewWebPart, _super);
    function HelloWorldNewWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        return _this;
    }
    HelloWorldNewWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n    <section class=\"".concat(styles.helloWorldNew, " ").concat(!!this.context.sdks.microsoftTeams ? styles.teams : '', "\">\n      <div class=\"").concat(styles.welcome, "\">\n        <img alt=\"\" src=\"").concat(this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png'), "\" class=\"").concat(styles.welcomeImage, "\" />\n        <h2>Well done, ").concat(escape(this.context.pageContext.user.displayName), "!</h2>\n        <div>").concat(this._environmentMessage, "</div>\n      </div>\n      <div>\n        <h3>Welcome to SharePoint Framework!</h3>\n        <div>Web part description: <strong>").concat(escape(this.properties.description), "</strong></div>\n        <div>Web part test: <strong>").concat(escape(this.properties.test), "</strong></div>\n        <div>Loading from: <strong>").concat(escape(this.context.pageContext.web.title), "</strong></div>\n      </div>\n      <div id=\"spListContainer\" />\n    </section>");
        this._renderListAsync();
    };
    HelloWorldNewWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
        });
    };
    HelloWorldNewWebPart.prototype._getEnvironmentMessage = function () {
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
    HelloWorldNewWebPart.prototype.onThemeChanged = function (currentTheme) {
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
    Object.defineProperty(HelloWorldNewWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    HelloWorldNewWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                PropertyPaneTextField('test', {
                                    label: 'Multi-line Text Field',
                                    multiline: true
                                }),
                                PropertyPaneCheckbox('test1', {
                                    text: 'Checkbox'
                                }),
                                PropertyPaneDropdown('test2', {
                                    label: 'Dropdown',
                                    options: [
                                        { key: '1', text: 'One' },
                                        { key: '2', text: 'Two' },
                                        { key: '3', text: 'Three' },
                                        { key: '4', text: 'Four' }
                                    ]
                                }),
                                PropertyPaneToggle('test3', {
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
    HelloWorldNewWebPart.prototype._getListData = function () {
        return this.context.spHttpClient.get("".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists?$filter=Hidden eq false"), SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        })
            .catch(function (err) { console.log(err); });
    };
    HelloWorldNewWebPart.prototype._renderList = function (items) {
        var html = '';
        items.forEach(function (item) {
            html += "\n    <ul class=\"".concat(styles.list, "\">\n      <li class=\"").concat(styles.listItem, "\">\n        <span class=\"ms-font-l\">").concat(item.Title, "</span>\n      </li>\n    </ul>");
        });
        if (this.domElement.querySelector('#spListContainer') !== null) {
            this.domElement.querySelector('#spListContainer').innerHTML = html;
        }
    };
    HelloWorldNewWebPart.prototype._renderListAsync = function () {
        var _this = this;
        this._getListData()
            .then(function (response) {
            _this._renderList(response.value);
        })
            .catch(function (err) { console.log(err); });
    };
    return HelloWorldNewWebPart;
}(BaseClientSideWebPart));
export default HelloWorldNewWebPart;
//# sourceMappingURL=HelloWorldNewWebPart.js.map