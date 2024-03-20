import { Version } from '@microsoft/sp-core-library';
/** CSJ Comment */
/** Property declaration step 01 */
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
/** CSJ Comment*/
/** BaseClientSideWebPart implements the minimal functionality that is required to build a web part.*/
/** This class also provides many parameters to validate and access read-only properties such as displayMode,
web part properties, web part context, web part instanceId, the web part domElement*/
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
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
    private _isDarkTheme;
    private _environmentMessage;
    /** CSJ Comment */
    /** The render() method is used to render the web part inside that DOM element.
       In the web part, the DOM element is set to a DIV. */
    /** Notice how ${ } is used to output the variable's value in the HTML block.
  * An extra HTML div is used to display this.context.pageContext.web.title. */
    render(): void;
    /** CSJ Comment */
    private _getListData;
    private _renderList;
    private _renderListAsync;
    protected onInit(): Promise<void>;
    private _getEnvironmentMessage;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected get dataVersion(): Version;
    /** CSJ Comment */
    /** The property pane is defined in the SecondExerciseNoJsFrameworkWebPart class.
        The getPropertyPaneConfiguration property is where you need to define the property pane. */
    /** When the properties are defined, you can access them in your web part
        by using this.properties.<property-value>, as shown in the render() method: */
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SecondExerciseWebPartWebPart.d.ts.map