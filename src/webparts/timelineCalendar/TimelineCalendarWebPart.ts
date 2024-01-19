import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneLabel, PropertyPaneToggle, PropertyPaneLink } 
  from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TimelineCalendarWebPartStrings';
import TimelineCalendar from './components/TimelineCalendar';
import { ITimelineCalendarProps } from './components/ITimelineCalendarProps';
import { ICalendarItem, IPersonaProps, IMemberOfResult, ICategoryItem, IGroupItem, IListItem } from './components/IConfigurationItems';

import { PropertyFieldCollectionData, CustomCollectionFieldType, ICustomDropdownOption, ICustomCollectionField } from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
//import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { PropertyFieldMonacoEditor } from '@pnp/spfx-property-controls/lib/PropertyFieldMonacoEditor';
//import { MonacoEditor?? } from "@pnp/spfx-controls-react/lib/MonacoEditor??";
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
//import { PropertyPaneMarkdownContent } from '@pnp/spfx-property-controls/lib/PropertyPaneMarkdownContent';
import { PropertyFieldMessage } from '@pnp/spfx-property-controls/lib/PropertyFieldMessage';
import { MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import PnPTelemetry from "@pnp/telemetry-js";
import MonacoPanelEditor from './components/MonacoPanelEditor';
import { PanelType, IDropdownOption, DropdownMenuItemType } from 'office-ui-fabric-react';
import AsyncDropdown from './components/AsyncDropdown';
import { MSGraphClientV3, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Guid } from '@microsoft/sp-core-library';
import { GraphError } from '@microsoft/microsoft-graph-client'; //ResponseType
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import DirectoryPicker from './components/DirectoryPicker';

//These are the persisted web part properties
export interface ITimelineCalendarWebPartProps {
  description: string;
  categories: any[];
  groups: any[];
  lists: any[];
  calsAndPlans: any[];
  minDays: number;
  maxDays: number;
  initialStartDays: number;
  initialEndDays: number;
  holidayCategories: string;
  fillFullWidth: boolean;
  calcMaxHeight: boolean;
  singleDayAsPoint: boolean;
  overflowTextVisible: boolean;
  hideItemBoxBorder: boolean;
  //hideSocialBar: boolean;
  //getDatesAsUtc: boolean;
  tooltipEditor: string;
  visJsonProperties: string;
  cssOverrides: string;
}

export default class TimelineCalendarWebPart extends BaseClientSideWebPart<ITimelineCalendarWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private dataCache = { //Used mostly to store REST promises for faster retrieval within PropertyFieldCollectionData
    webs: {} as any,
    lists: {} as any,
    views: {} as any,
    fields: {} as any,
    memberOf: [] as IPersonaProps[],
    calendars: {} as any,
    calFilterQuery: {} as any
  };
  private _graphClient: Promise<MSGraphClientV3> = null;
  private _graphScopes: string[] = [];
  private _messageListener = (event:MessageEvent) => {
    //Look for only "local" events (ignoring OWA webshell messages: https://webshell.dodsuite.office365.us)
    if (event.origin === window.location.origin) {
      //event.data == 
        //Property pane will open (also when *updating props* and when switching to other web parts, but doesn't cause following "toggled" to fire)
        //Property pane toggled (when opened and also when closed, also when page is saved)
      //event.timeStamp: 219218.10000002384 (will be the same for duplicated messages)
      if (event.data == "Property pane toggled")
        //Reset the cache
        this.dataCache = {
          webs: {},
          lists: {},
          views: {},
          fields: {},
          memberOf: this.dataCache.memberOf,
          calendars: {},
          calFilterQuery: {}
        }
    }
  };
  
  //NOTE: This is fired only once once the web part is initially loading
  //But when page is *edited* it is fired again (displayMode == 2 in this case)
  //Also fired while already in edit mode and new web part is added (also mode #2)
  protected onInit(): Promise<void> {
    //console.log("onInit, displayMode:" + this.displayMode.toString());
    if (this.displayMode == DisplayMode.Edit)
      window.addEventListener("message", this._messageListener, false);

    //Opt-out of PnP telemetry
    const telemetry = PnPTelemetry.getInstance();
    telemetry.optOut();

    //Get reference
    this._graphClient = this.context.msGraphClientFactory.getClient('3');
    
    //Just while in dev testing mode
    /*if (this.context.isServedFromLocalhost) { //because onDisplayModeChanged is not changed
      console.log("isServedFromLocalhost");
      this.getUserMemberOf();
    }*/

    //Get user's groups in case Outlook Calendars collection/pane is opened
    if (this.displayMode == DisplayMode.Edit)
      this.getUserMemberOf();

    //If there's no existing data, add some default categories and groups to give the user a visual starting point/example
    if (this.properties.categories == null)
      this.properties.categories = [
        {
          uniqueId: Guid.newGuid().toString(),
          name: 'Category 1',
          borderColor: '#06a303',
          bgColor: '#4cfc4c',
          textColor: '#1a1a1a',
          visible: true,
          sortIdx: 1,
          advancedStyles: null
        } as ICategoryItem,
        {
          uniqueId: Guid.newGuid().toString(),
          name: 'Category 2',
          borderColor: '#d9a302',
          bgColor: '#ffe28a',
          textColor: '#1a1a1a',
          visible: true,
          sortIdx: 2,
          advancedStyles: null
        } as ICategoryItem,
        {
          uniqueId: Guid.newGuid().toString(),
          name: 'Category 3',
          borderColor: '#b30707',
          bgColor: '#fa8e8e',
          textColor: '#1a1a1a',
          visible: true,
          sortIdx: 3,
          advancedStyles: null
        } as ICategoryItem
      ];

    if (this.properties.groups == null)
      this.properties.groups = [
        {
          uniqueId: Guid.newGuid().toString(),
          name: 'Row 1',
          visible: true,
          html: null,
          sortIdx: 1
        } as IGroupItem,
        {
          uniqueId: Guid.newGuid().toString(),
          name: 'Row 2',
          visible: true,
          html: null,
          sortIdx: 2
        } as IGroupItem,
        {
          uniqueId: Guid.newGuid().toString(),
          name: 'Row 3',
          visible: true,
          html: null,
          sortIdx: 3
        } as IGroupItem
      ];

    //Set default value for tooltip editor
    if (this.properties.tooltipEditor == null || this.properties.tooltipEditor == "")
      this.properties.tooltipEditor = this.getDefaultTooltip();

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  //NOTE: This is fired even for property changes made in the properties pane! (props pane is this.displayMode == 2)
  //this.context (and .instanceId) is valid here
  public render(): void {
    //console.log("render, displayMode:" + this.displayMode.toString());
    const element: React.ReactElement<ITimelineCalendarProps> = React.createElement(TimelineCalendar,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        instanceId: this.context.instanceId,
        categories: this.properties.categories,
        groups: this.properties.groups,
        lists: this.properties.lists,
        calsAndPlans: this.properties.calsAndPlans,
        //renderLegend: this.renderLegend.bind(this), //called from TSX, .bind needed otherwise "this" refers to the .tsx
        getDefaultTooltip: this.getDefaultTooltip.bind(this),
        ensureValidClassName: this.ensureValidClassName.bind(this),
        buildDivStyles: this.buildDivStyles.bind(this),
        context: this.context,
        graphClient: this._graphClient,
        getGraphScopes: this.getGraphScopes.bind(this),
        domElement: this.domElement,
        minDays: this.properties.minDays,
        maxDays: this.properties.maxDays,
        initialStartDays: this.properties.initialStartDays,
        initialEndDays: this.properties.initialEndDays,
        holidayCategories: this.properties.holidayCategories,
        fillFullWidth: this.properties.fillFullWidth,
        calcMaxHeight: this.properties.calcMaxHeight,
        singleDayAsPoint: this.properties.singleDayAsPoint,
        overflowTextVisible: this.properties.overflowTextVisible,
        hideItemBoxBorder: this.properties.hideItemBoxBorder,
        //hideSocialBar: this.properties.hideSocialBar,
        //getDatesAsUtc: this.properties.getDatesAsUtc,
        tooltipEditor: this.properties.tooltipEditor,
        visJsonProperties: this.properties.visJsonProperties,
        cssOverrides: this.properties.cssOverrides
      }
    );
    ReactDom.render(element, this.domElement);
  }
  
  private getUserMemberOf():void {
    if (this.displayMode == DisplayMode.Edit) {
      this._graphClient.then((client:MSGraphClientV3): void => {
        client.api("/me/memberOf").select("id,displayName,mail,visibility")
        //Need to exclude security groups
        .filter("groupTypes/any(c:c eq 'Unified')") //filter could also be "mailEnabled eq true" with &$count=true
        .header('ConsistencyLevel', 'eventual').count() //both needed for filter to work
        .get((error:GraphError, response:any, rawResponse?:any) => {
          //Handle errors
          if (error) {
            //Nothing needed
            console.error(error);
          }
          //Handle a success response
          else {
            if (response == null) {
              this.dataCache.memberOf = [];
            }
            else {
              //TODO: https://stackoverflow.com/questions/24806772/how-to-skip-over-an-element-in-map
              this.dataCache.memberOf = response.value.filter((value:IMemberOfResult, index:number) => {
                //return (value['@odata.type'] == '#microsoft.graph.group');
                return (value.visibility != null);
              }).map((value:IMemberOfResult, index:number):IPersonaProps => {
                return {
                  key: value.id,
                  //imageInitials: '',
                  mail: value.mail,
                  text: value.displayName,
                  //secondaryText: value.visibility + " group", //TODO: Third visbility type of hiddenmembership (so show only Public or Private)
                  secondaryText: value.mail,
                  personaType: "group"
                }
              });

              //const groups:IMemberOfResult[] = response.value;
              //const aGroup = groups[0];

              //@odata.type: "#microsoft.graph.directoryRole"
              //displayName: null
              //@odata.type: "#microsoft.graph.group"
              //console.log(aGroup.displayName);
            }
          }
        });
      });
    }
  }
  
  //NOTE: This is fired before any onChange function from properties (i.e. MonacoEditor)
  protected override onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    //Special handling for some properties
    if (propertyPath === "groups") {
      //Check if there were no groups but now new ones added; need to tag any lists to a group or their events won't show
      if ((oldValue == null || oldValue.length == 0) && (newValue && newValue.length > 0) && this.properties.lists) {
        const categoryId = (this.properties.groups[0] as IGroupItem).uniqueId; //get id for first group
        this.properties.lists.forEach((list: IListItem) => {
          list.group = "Static:" + categoryId;
        })
      }
    }
    else if (propertyPath === "visJsonProperties") {
      //Proceed saving the data only if it's valid JSON
      try {
        JSON.parse(newValue);
      }
      catch (e) {
        this.properties.visJsonProperties = oldValue; //Overwrite back to original value
      }
    }

    //After this the render() function is fired followed by componentDidUpdate() in the .tsx
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);

    //console.log("onDispose, displayMode:" + this.displayMode.toString());
    if (this.displayMode == DisplayMode.Edit)
      //Remove event listener
      window.removeEventListener("message", this._messageListener, false);

    //Remove the styles that were dynamically added
    const styleId = "TimelineDynStyles-" + this.instanceId.substring(24); //use last portion of GUID
    const styleElem = document.getElementById(styleId);
    if (styleElem)
      styleElem.parentNode.removeChild(styleElem);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private getDefaultTooltip(): string {
    return `<!-- Two curly braces indicates a property/field value will be injected -->
<div class="vis-tooltipTitle">{{content}}</div>
<div class="vis-tooltipBody">
  <div><b>Location:</b> {{Location}}</div>
  <div><b>Category:</b> {{Category}}</div>
  <!-- "date" is a handler to format a date value -->
  <div><b>Start:</b> {{date start}}</div>
  <div><b>End:</b> {{date end}}</div>
  <!-- Using three curly braces *disables* HTML character escaping -->
  <!-- "limit" is a handler to shorten *potentially* long text -->
  <div><b>Description:</b> {{{limit Description}}}</div>
  <hr />
  <div style='font-size:0.9em;'>
    <div><b>Created By:</b> {{Author}}</div>
    <div><b>Modified By:</b> {{Editor}}</div>
    <div><b>Modified On:</b> {{date Modified}}</div>
  </div>
</div>`;
  }

  private ensureValidClassName(className:string | []): string {
    if (className == null || className.length == 0)
      return null; //or "" or cal.className ???

    if (Array.isArray(className))
      className = (className as []).join(", ");

    className = (className as string); //just for TypeScript compiling

    //Calculated fields add extra content, remove it
    if (className.indexOf(";#") != -1) { //ex: string;#CalculatedValueHere
      const index = className.indexOf(";#");
      className = className.substring(index+2);
    }
    
    //Ensure valid CSS classes (no spaces, reserved characters, etc.)
    className = className.replace(/\W/g, "");
    
    //Check if class starts with a number, which isn't valid
    if (isNaN(Number(className.charAt(0))) == false)
      //className = TC.settings.numCssClassPrepend + className;
      className = "Prepend" + className;
    return className;
  }

  private buildDivStyles(categoryItem:ICategoryItem): any {
    const defaultStyle = "border-color:" + categoryItem.borderColor + "; color:" + categoryItem.textColor + ";" + 
      (categoryItem.bgColor ? " background-color:" + categoryItem.bgColor + ";" : "");

    if (categoryItem.advancedStyles) {
      //Find all content within {} characters (the styles within the defined CSS class)
      //categoryItem.advancedStyles.match(/\{([^{]+)\}/g); //global returns a single "concatinated" array result
      const stylesMatch = categoryItem.advancedStyles.match(/\{([^{]+)\}/); //without global the second array result doesn't have {} characters
      if (stylesMatch && stylesMatch[1])
        return stylesMatch[1];
      else
        return defaultStyle;
    }
    else
      return defaultStyle;
  }

  private getGraphScopes(returnError?:boolean): string[] {
    //Look for and save Graph scope information
    for (const key in sessionStorage) {
      //@ts-ignore (for startsWith)
      if (key && typeof key == "string" && key.startsWith('{"authority":')) {
        try {
          const keyObj = JSON.parse(key);
          //Find the correct Graph results entry
          //TODO: Look for the SPO client app specifically (decode the JWT)
          if (keyObj.scopes && keyObj.scopes.indexOf('profile openid') != -1) {
            const scopes = keyObj.scopes.split(" ");
            //Change "https://[dod-]graph.microsoft.[com/us]/User.Read" value to instead just get "User.Read" portion
            this._graphScopes = scopes.map((value:string, i:Number) => {
              if (value.indexOf('graph.microsoft.') != -1)
                return value.split("/")[3];
              else
                return value;
            //Remove the "profile", "openid", "email", and ".default" scopes
            }).filter((value:string) => {
              switch (value) {
                case "profile":
                case "openid":
                case "email":
                case ".default": //this one has the Graph URL prepended to it, which is why this .filter is after .map
                  return false;
                
                default:
                  return true;
              }
            });
          }
        }
        catch (e) {}
      }
    }

    //If returnError specified, return the error details if there were no scopes found
    if ((this._graphScopes == null || this._graphScopes.length == 0) && returnError 
          && sessionStorage["msal.error.description"])
      this._graphScopes = [sessionStorage["msal.error.description"]];

    return this._graphScopes;
  }

  //Fired each time property pane is opened (initial and close-open actions)
  //Also fired after render() after property pane properties are saved/changed
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    //Save references because "this" is not available  within the "return" below
    const pageContext = this.context.pageContext;
    const spHttpClient = this.context.spHttpClient;
    const self = this;

    //MarkDown for web part information (make sure to remove left indentation/spaces)
    /*const webpartMD = `**Web Part Version**

${this && this.manifest.version ? this.manifest.version : '*Unknown*'}

**Web Part Instance ID**

${this.instanceId}
`;*/

    //Return the PropertyPane config
    return {
      pages: [
        {
          /*header: {
            description: strings.PropertyPaneDescription
          },*/
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldMessage("", {
                  key: "topInstructions",
                  //NOTE: Use of HTML is escaped :(
                  text: "For help using this web part, refer to the last page (bottom right, click Next) to find the support and documentation links",
                  messageType: MessageBarType.info,
                  isVisible: true
                }),
                //Reference: https://pnp.github.io/sp-dev-fx-property-controls/controls/PropertyFieldCollectionData/
                PropertyFieldCollectionData("categories", {
                  key: "categories",
                  value: this.properties.categories,
                  label: "Categories / Legend", //Header/Label above the button
                  manageBtnLabel: "Add/Edit Categories",
                  panelProps: {
                    type: PanelType.largeFixed, //.large is default, but .largeFixed works ok for this one
                    //customWidth: "820px" //was 750px when background was not included
                  },
                  panelHeader: "Configure Categories / Legend",
                  panelDescription: "Categories defined here appear as filterable legend items above the timeline.",
                  saveBtnLabel: "Save & close",
                  saveAndAddBtnLabel: "Add/Save & close",
                  enableSorting: true,
                  fields: [
                    {
                      id: "name",
                      title: "Category Name",
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: " " //need a space because blank just shows the title
                    },
                    {
                      id: "borderColor",
                      title: "Border", //was "Color" when background not shown
                      type: CustomCollectionFieldType.color,
                      defaultValue: '#97b0f8',
                      required: false
                    },
                    {
                      id: "bgColor",
                      title: "Background",
                      type: CustomCollectionFieldType.color,
                      defaultValue: '#d5ddf6',
                      required: false
                    },
                    {
                      id: "textColor",
                      title: "Text",
                      type: CustomCollectionFieldType.color,
                      defaultValue: '#1a1a1a',
                      required: false,
                      //Oddly named: This is really the "perform field validation" function
                      onGetErrorMessage: (value: string, index: number, item: ICategoryItem) => {
                        //Using this function just to override the colors in the new/bottom row
                        //When sortIdx doesn't exists, it's a new row
                        if (item.sortIdx == null) { //item.name should have a value because this field was just edited
                          //Update actual field values to match what the color fields show (usually the same as last entered row)
                          document.querySelectorAll("div.ms-Panel-content div.PropertyFieldCollectionData__panel__table-row:last-child div.PropertyFieldCollectionData__panel__color-field").forEach(function(elem, index) {
                            switch (index) {
                              case 0: //Border
                                item.borderColor = (elem.children[0] as HTMLDivElement).style.backgroundColor;
                                break;
                              case 1: //Background
                                item.bgColor = (elem.children[0] as HTMLDivElement).style.backgroundColor;
                                break;
                              case 2: //Text
                                item.textColor = (elem.children[0] as HTMLDivElement).style.backgroundColor;
                                break;
                            }
                          });
                        }
                        //Need to let default validation happen
                        return '';
                      },
                    },
                    /*{
                      id: "fixColors",
                      title: "fixColors",
                      isVisible: (field, items) => {
                        return false;
                      },
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item:ICategoryItem, itemId, onCustomFieldValidation) => {
                        //const divStyles = { backgroundColor: item.Background, borderColor: item.Border, color:item.Text};
                        //or https://stackoverflow.com/questions/56934170/creating-styled-component-with-react-createelement
                        
                        //When sortIdx exists it's an existing row (not a new one)
                        if (item.sortIdx == null && item.name == null) { //When item.name == "" it had a value but was cleared out (even for new rows)
                          //This is an existing row which can be handled normally
                          //return React.createElement("span", {class: 'legendBox vis-item vis-range', style: { //padding: "0 5px", borderRadius: "2px",
                          //  borderColor: item.borderColor, color:item.textColor}}, item.name); //border: "1px solid " + item.borderColor
                          
                          //Clear color boxes to handle bug where the last added row's color fields copy to
                          //this new one and it doesn't report the actual colors in the "item" object
                          document.querySelectorAll("div.ms-Panel-content div.PropertyFieldCollectionData__panel__table-row:last-child div.PropertyFieldCollectionData__panel__color-field").forEach(function(elem, index) {
                            switch (index) {
                              case 0: //Border
                                (elem.children[0] as HTMLDivElement).style.backgroundColor = "#97b0f8"; //Default vis.js colors
                                break;
                              case 1: //Background
                                (elem.children[0] as HTMLDivElement).style.backgroundColor = "transparent";
                                break;
                              case 2: //Text
                                (elem.children[0] as HTMLDivElement).style.backgroundColor = "#1a1a1a";
                                break;
                            }
                          });
                        }
                        if (item.name == null || item.name == "") //Don't render a styled box at all
                          return React.createElement("span", null, null);
                        else {
                          const divStyles = self.buildDivStyles(item);
                          return React.createElement("span", {class: 'legendBox vis-item vis-range', style: divStyles}, item.name);
                        }
                      }
                    },*/
                    {
                      id: "visible",
                      title: "Visible",
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: true,
                    },
                    {
                      id: "advancedStyles",  
                      title: "Advanced styles",
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item:ICategoryItem, itemId, onCustomFieldValidation) => {  
                        return (  
                          React.createElement(MonacoPanelEditor, {
                            key: itemId,
                            disabled: false,
                            buttonText: "Advanced Edit",
                            headerText: 'Advanced editor for Category styles',
                            value: (value || "/* Customize your desired styles below */\r\n.CategoryClassName {\r\n  border-color: " + item.borderColor + ";\r\n  background-color: " + item.bgColor + ";\r\n  color: " + item.textColor + ";\r\n}"),
                            language: "css",
                            onValueChanged: (newValue: string) => {
                              onUpdate(field.id, newValue);
                            }
                          })
                        )
                      }
                    }
                  ],
                }),
                PropertyFieldCollectionData("groups", {
                  key: "groups",
                  value: this.properties.groups,
                  label: "Rows / Swimlanes", //Header/Label above the button
                  manageBtnLabel: "Add/Edit Rows",
                  panelProps: {
                    type: PanelType.largeFixed //was PanelType.medium before CSS
                  },
                  panelHeader: "Configure Rows / Swimlanes",
                  panelDescription: "Rows defined here appear as horizontal swimlanes in the timeline.",
                  saveBtnLabel: "Save & close",
                  saveAndAddBtnLabel: "Add/Save & close",
                  enableSorting: true,
                  fields: [
                    {
                      id: "name",
                      title: "Row name",
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: " " //need a space because blank just shows the title
                    },
                    {
                      id: "className",
                      title: "CSS Class",
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: " ", //need a space because blank just shows the title
                      //deferredValidationTime: 1000,
                      //Oddly named: This is really the "perform field validation" function
                      onGetErrorMessage: (value: string, index: number, item: IGroupItem) => {
                        //NOTE: "this" is just the field object
                        //Fired after deferredValidationTime
                        
                        //Handle blank and cleared-out values
                        if (value == null || value.trim() == '') {
                          //Nothing
                        } else
                          //Force into a valid CSS class name
                          item.className = self.ensureValidClassName(value);
                        
                        //Always return, or Save button is disabled
                        return ''; //no validation error; '' lets default checks happen
                      }
                    },
                    {
                      id: "visible",
                      title: "Visible",
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: true
                    },
                    {  
                      id: "html", //TODO: Consider https://sharepoint.stackexchange.com/questions/277786/retrieving-data-from-rich-text-editor-in-spfx-web-part-properties
                      title: "Advanced HTML",  
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item:IGroupItem, itemId, onCustomFieldValidation) => {  
                        return (  
                          React.createElement(MonacoPanelEditor, {
                            key: itemId,
                            disabled: false,
                            buttonText: "Advanced Edit",
                            headerText: 'Row / Swimlane HTML Content',
                            value: (value || "<!-- Two curly braces indicates a property/field value will be injected -->\r\n<div>{{name}}</div>"), //" + item.name + "
                            language: "html",
                            onValueChanged: (newValue: string) => {
                              onUpdate(field.id, newValue);
                            }
                          })
                        )
                      }
                    }
                  ],
                })
              ]
            },
            {
              groupName: "Data Source Settings",
              groupFields: [
                PropertyFieldCollectionData("lists", {
                  key: "lists",
                  value: this.properties.lists,
                  label: "SharePoint Lists / Calendars", //Header/label above the button
                  manageBtnLabel: "Add/Edit Lists",
                  panelProps: {
                    type: PanelType.smallFluid
                  },
                  panelHeader: "Configure SharePoint Lists / Calendars",
                  panelDescription: "Specify the desired lists and, optionally, the views to use and category option.",
                  saveBtnLabel: "Save & close",
                  saveAndAddBtnLabel: "Add/Save & close",
                  //enableSorting: true, //not necessary here as the list order doesn't matter
                  fields: [
                    {
                      id: "siteUrl",
                      title: "Site URL",
                      defaultValue: pageContext.web.serverRelativeUrl, //current site
                      type: CustomCollectionFieldType.string,
                      required: true,
                      deferredValidationTime: 1000,
                      //Oddly named: This is really the "perform field validation" function
                      onGetErrorMessage: (value: string, index: number, item: IListItem) => {
                        //NOTE: "this" is just the field object
                        //Fired after deferredValidationTime
                        
                        //Force reset the List field to clear out any previously selected value if the Site is changed
                        //item.List = null; //Cannot do as this func gets fired after a List selection too

                        //Handle blank and cleared-out values
                        if (value == null || value.trim() == '')
                          return ''; //'' lets default checks happen w/o showing red border (will still show 'Site is required' in red warning circle icon)

                        //Was a full path/URL entered (https://...)
                        if (value.trim().substring(0, 1) != "/") {
                          let rootUrl = pageContext.web.absoluteUrl;
                          if (pageContext.web.serverRelativeUrl !== "/")
                            //Get just the root domain, ex: https://usaf.dps.mil/
                            rootUrl = pageContext.web.absoluteUrl.replace(pageContext.web.serverRelativeUrl, '');
                          
                          //Compare to user entered domain value
                          if (value.indexOf(rootUrl) != 0)
                            return 'Site must be on the same domain';
                          else //Make the URL relative
                            value = value.replace(rootUrl, "");
                        }

                        //Remove any trailing slash
                        if (value.lastIndexOf("/") + 1 == value.length)
                        value = value.substring(0, value.length-1);

                        //Get just the base site URL if these known URL formats were provided
                        value = value.split("/Lists/")[0];
                        value = value.split("/lists/")[0];
                        value = value.split("/Pages/")[0];
                        value = value.split("/pages/")[0];
                        value = value.split("/SitePages/")[0];
                        value = value.split("/sitepages/")[0];
                        //TODO for library: https://usaf.dps.mil/teams/UA-App-VCP/csktest/ODSTest/Forms/AllItems.aspx

                        //Look if .aspx is still at end of the URL to warn user
                        //@ts-ignore (we know endsWith is available)
                        if (value.toLowerCase().endsWith(".aspx"))
                          return 'URL must be to the site only and not to a list or page';
                        
                        //Update the field value with the shortened/processed URL
                        item.siteUrl = value;

                        //Look for existing web check and return it's promise
                        if (self.dataCache.webs[value])
                          return self.dataCache.webs[value]

                        //URL should be structured correctly at this point, but may not be a valid site
                        const promise = new Promise<string>((resolve, reject) => {
                          spHttpClient.get(value + "/_api/web?$select=Id", SPHttpClient.configurations.v1) //or Title,ServerRelativeUrl
                            .then((response: SPHttpClientResponse) => {
                              if (response.status == 404) {
                                resolve("Could not resolve site, please verify.");
                                return;
                              }

                              if (response.ok)
                                resolve(''); //no validation error
                              else {
                                const statusCode = response.status;
                                //const statusMessage = response.statusMessage; //says doesn't exist
                                response.json().then(data => {
                                  console.log(data);
                                  resolve(data.error.message);
                                })
                                .catch (error => {
                                  //resolve(error.message);
                                  resolve("Error HTTP: " + statusCode.toString() + " " + response.statusText);
                                  //Reset the promise cache in case of temp issue
                                  self.dataCache.webs[value] = null;
                                });
                              }
                            })
                            .catch(error => {
                              //console.log(error);
                              /*
                              .message: "Unexpected end of JSON input"
                              .stack: "SyntaxError: Unexpected end of JSON input\n    at e.json..."
                              */
                              resolve(error.message);
                              //Reset the promise cache in case of temp issue
                              self.dataCache.webs[value] = null;
                            });
                        });

                        //Store the promise in cache and return
                        self.dataCache.webs[value] = promise;
                        return promise;
                      }
                    },
                    {  
                      id: "list",
                      title: "List",  
                      required: true,
                      //onGetErrorMessage(value:any, index:number, currentItem:any) { //only fires when value changed (not at initial creation)...
                      //...use onCustomFieldValidation within onCustomRender
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately after Site field change; not honoring deferredValidationTime on Site field
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {  
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: item.siteUrl,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null)
                                onUpdate(field.id, null);
                              else {
                                //ListTemplateType: https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee541191(v=office.15)
                                if (option.data.baseTemplate == 106) {
                                  item.isCalendar = true;
                                  item.startDateField = "EventDate"; //Set these known values for the user
                                  item.endDateField = "EndDate";
                                  item.category = "Field:Category";
                                }
                                else {
                                  item.isCalendar = false;
                                  item.startDateField = null;
                                  item.endDateField = null;
                                  item.category = null;
                                }
                                
                                //Finalize the change
                                item.listName = option.text;
                                onUpdate(field.id, option.key);
                              }
                            },
                            loadOptions: () => {
                              //NOTE: "this" is TimelineCalendarWebPart with .render, properties
                              //field (id:list, required: true, title: "List") and item are available

                              //Don't really need this since it's the first selection that enables the other fields
                              //if (value == null)
                              //  onCustomFieldValidation(field.id, ''); //let the default "required" message show

                              //Look for an existing lists check and return it's promise
                              if (self.dataCache.lists[item.siteUrl])
                                return self.dataCache.lists[item.siteUrl];

                              //Get non-catalog, "regular" lists from the site
                              const listPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                //Was including "and BaseTemplate le 106" but that cut out CustomGrid lists created by exporting Excel data
                                //https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wssts/8bf797af-288c-4a1d-a14b-cf5394e636cf
                                //https://github.com/pnp/pnpcore/blob/dev/src/sdk/PnP.Core/Model/SharePoint/Core/Public/Enums/ListTemplateType.cs
                                //RootFolder/Name == "Workflow History"
                                //RootFolder/ServerRelativeUrl == "/sites/TipsToolsApps/Lists/Workflow History"
                                spHttpClient.get(item.siteUrl + "/_api/web/lists?$select=BaseTemplate,BaseType,Id,Hidden,Title,RootFolder/ServerRelativeUrl" +
                                  "&$expand=RootFolder&$filter=IsCatalog eq false and IsPrivate eq false", SPHttpClient.configurations.v1)
                                  .then((response: SPHttpClientResponse) => {
                                    if (response.ok) {
                                      //TODO: Instead call .text() and then try/catch with JSON.parse?
                                      response.json().then((data:any) => {
                                        let promiseData:IDropdownOption[] = [];
                                        let calendars:IDropdownOption[] = [];
                                        let lists:IDropdownOption[] = [];
                                        let others:IDropdownOption[] = [];
                                        data.value.forEach((list:any) => {
                                          //Check for "legacy" calendars
                                          if (list.BaseTemplate == 106)
                                            calendars.push({
                                              key: list.Id,
                                              text: list.Title,
                                              data: {
                                                baseTemplate: list.BaseTemplate
                                              }
                                            })
                                          else {
                                            const ignoredLists = [
                                              {
                                                BaseTemplate: 160,
                                                Hidden: true,
                                                WebRelativeUrl: "/Access Requests"
                                              },
                                              {
                                                BaseTemplate: 125,
                                                Hidden: true,
                                                WebRelativeUrl: "/_catalogs/appdata"
                                              },
                                              {
                                                BaseTemplate: 336,
                                                WebRelativeUrl: "/AppCatalog"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Cache Profiles"
                                              },
                                              {
                                                BaseTemplate:334,
                                                Hidden: true,
                                                WebRelativeUrl: "/ClientSideAssets"
                                              },
                                              {
                                                BaseTemplate: 331,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/ComponentManifests"
                                              },
                                              {
                                                BaseTemplate: 124,
                                                Hidden: true,
                                                WebRelativeUrl: "/_catalogs/design"
                                              },
                                              {
                                                //Title: "Content and Structure Reports",
                                                BaseTemplate: 100,
                                                WebRelativeUrl: "/Reports List"
                                              },
                                              {
                                                //Title: "Content Organizer Rules"
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/RoutingRules"
                                              },
                                              {
                                                //Title: "Content type service application error log",
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/ContentTypeAppLog"
                                              },
                                              {
                                                //Title: Content type publishing error log
                                                BaseTemplate:100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/ContentTypeSyncLog"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/DeviceChannels"
                                              },
                                              {
                                                //Title: "DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES"
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CON"
                                              },
                                              {
                                                BaseTemplate: 101,
                                                WebRelativeUrl: "/FormServerTemplates"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Long Running Operation Status"
                                              },
                                              {
                                                //Title: "MicroFeed"
                                                BaseTemplate: 544,
                                                WebRelativeUrl: "/Lists/PublishedFeed"
                                              },
                                              {
                                                //Title: Notification List
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Notification Pages"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/ProjectPolicyItemList"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Quick Deploy Items"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Relationships List"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/Reporting Metadata"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                WebRelativeUrl: "/ReusableContent"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/SharePointHomeCacheList"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/PackageList"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/PageDiagnosticsResultList70B82488D89F4AAPDR"
                                              },
                                              {
                                                BaseTemplate: 3300,
                                                Hidden: true,
                                                WebRelativeUrl: "/Sharing Links"
                                              },
                                              {
                                                BaseTemplate: 101,
                                                WebRelativeUrl: "/SiteAssets"
                                              },
                                              {
                                                BaseTemplate: 101,
                                                WebRelativeUrl: "/SiteCollectionDocuments"
                                              },
                                              {
                                                BaseTemplate: 851,
                                                WebRelativeUrl: "/SiteCollectionImages"
                                              },
                                              {
                                                BaseTemplate: 119,
                                                WebRelativeUrl: "/SitePages"
                                              },
                                              {
                                                //Title: Suggested Content Browser Locations
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/PublishedLinks"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/TaxonomyHiddenList"
                                              },
                                              {
                                                BaseTemplate: 337,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/TenantWideExtensions"
                                              },
                                              {
                                                BaseTemplate: 101,
                                                Hidden: true,
                                                WebRelativeUrl: "/Translation Packages"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Translation Status"
                                              },
                                              {
                                                //Title: User Information List
                                                BaseTemplate: 112,
                                                Hidden: true,
                                                WebRelativeUrl: "/_catalogs/users"
                                              },
                                              {
                                                BaseTemplate: 100,
                                                Hidden: true,
                                                WebRelativeUrl: "/Variation Labels"
                                              },
                                              {
                                                //Title: Web Template Extensions
                                                BaseTemplate: 3415,
                                                Hidden: true,
                                                WebRelativeUrl: "/_catalogs/wte"
                                              },
                                              {
                                                BaseTemplate: 122,
                                                Hidden: true,
                                                WebRelativeUrl: "/_catalogs/wfpub"
                                              },
                                              {
                                                BaseTemplate: 140,
                                                Hidden: true,
                                                WebRelativeUrl: "/Lists/Workflow History"
                                              },
                                              {
                                                BaseTemplate: 107,
                                                WebRelativeUrl: "/WorkflowTasks"
                                              }
                                            ]

                                            //Ignore known lists
                                            const foundList = ignoredLists.filter(ignList => {
                                              return (ignList.BaseTemplate == list.BaseTemplate && list.RootFolder.ServerRelativeUrl.endsWith(ignList.WebRelativeUrl))
                                            })
                                            if (foundList.length > 0)
                                              return; //skip this list                                     

                                            //Look for lists
                                            if (list.BaseTemplate == 100 || list.BaseTemplate == 104 || list.BaseTemplate == 107 || list.BaseTemplate == 120 || list.BaseTemplate == 150 || list.BaseTemplate == 171 || list.BaseTemplate == 1100)
                                              lists.push({
                                                key: list.Id,
                                                text: list.Title,
                                                data: {
                                                  baseTemplate: list.BaseTemplate
                                                }
                                              })
                                            else //Everything else (libraries+)
                                              others.push({
                                                key: list.Id,
                                                text: list.Title,
                                                data: {
                                                  baseTemplate: list.BaseTemplate
                                                }
                                              })
                                          }
                                        });

                                        //Add "Calendars" header
                                        if (calendars.length > 0) {
                                          promiseData.push({
                                            key: "calendarsHeader",
                                            text: "Calendars",
                                            itemType: DropdownMenuItemType.Header
                                          });
                                          promiseData = promiseData.concat(calendars);
                                        }

                                        //Add "Lists" header
                                        if (lists.length > 0) {
                                          promiseData.push({
                                            key: "listsHeader",
                                            text: "Lists",
                                            itemType: DropdownMenuItemType.Header
                                          });
                                          promiseData = promiseData.concat(lists);
                                        }

                                        //Add "Others" header
                                        if (others.length > 0) {
                                          promiseData.push({
                                            key: "othersHeader",
                                            text: "Others",
                                            itemType: DropdownMenuItemType.Header
                                          });
                                          promiseData = promiseData.concat(others);
                                        }

                                        resolve(promiseData);
                                      }) //response.json().then
                                    } //response.ok
                                    else {
                                      //const statusCode = response.status;
                                      //const statusMessage = response.statusMessage; //May not exist?
                                      response.json().then(data => {
                                        console.log(data);
                                        reject(data.error.message);
                                      })
                                      .catch (error => {
                                        //console.log("status: " + statusCode.toString() + " / " + statusNum.toString());
                                        //reject(error.message);
                                        reject("Error HTTP: " + response.status.toString() + " " + response.statusText);
                                        //Reset the promise cache in case of temp issue
                                        self.dataCache.lists[item.siteUrl] = null;
                                      });
                                    }
                                  }) //spHttpClient.get().then
                                  .catch(error => {
                                    //console.log(error);
                                    /*
                                    .message: "Unexpected end of JSON input"
                                    .stack: "SyntaxError: Unexpected end of JSON input\n    at e.json..."
                                    */
                                    reject(error.message);
                                    //Reset the promise cache in case of temp issue
                                    self.dataCache.lists[item.siteUrl] = null;
                                  });
                              }); //listPromise = new Promise

                              //Store promise in cache and return
                              self.dataCache.lists[item.siteUrl] = listPromise;
                              return listPromise;
                            }
                          })
                        )
                      }
                    },
                    {
                      id: "view",  
                      title: "View",
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately; not honoring deferredValidationTime on other fields
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: item.list,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null || option.key == "") {
                                if (item.viewName)
                                  item.viewName = null; //clear it
                                if (item.viewFilter)
                                  item.viewFilter = null; //clear it

                                //Update field
                                onUpdate(field.id, null);
                              }
                              else {
                                item.viewName = option.text;
                                onUpdate(field.id, option.key);
                              }
                            },
                            loadOptions: () => {
                              //Look for an existing views check and return it's promise
                              if (self.dataCache.views[item.list])
                                return self.dataCache.views[item.list];

                              //Get non-personal (public) and non-hidden views
                              const viewPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                spHttpClient.get(item.siteUrl + "/_api/web/lists('" + item.list + "')/views?$select=BaseViewId,Id,ServerRelativeUrl,Title,ViewType,ViewQuery&$filter=PersonalView ne true and Hidden ne true", SPHttpClient.configurations.v1)
                                  .then((response: SPHttpClientResponse) => {
                                    if (response.ok) {
                                      response.json().then((data:any) => {
                                        //Add a blank option to *not* select a View
                                        let promiseData:IDropdownOption[] = [];
                                        promiseData.push({
                                          key: "", //blank
                                          text: ""
                                        });
                                        //Add results to dropdown
                                        data.value.forEach((view:any) => {
                                          if (!view.ServerRelativeUrl.endsWith("/calendar.aspx") && //or ViewQuery: <Where><DateRangesOverlap><FieldRef Name=\"EventDate\" /><FieldRef Name=\"EndDate\" /><FieldRef Name=\"RecurrenceID\" /><Value Type=\"DateTime\"><Month /></Value></DateRangesOverlap></Where>
                                                !view.ServerRelativeUrl.endsWith("/MyItems.aspx")) { //or Title:"Current Events", ViewQuery: "<Where><DateRangesOverlap><FieldRef Name=\"EventDate\" /><FieldRef Name=\"EndDate\" /><FieldRef Name=\"RecurrenceID\" /><Value Type=\"DateTime\"><Now /></Value></DateRangesOverlap></Where><OrderBy><FieldRef Name=\"EventDate\" /></OrderBy>
                                            //Ignore the "Calendar" and "Current Events" views since they limit the items returned
                                            promiseData.push({
                                              key: view.Id,
                                              text: view.Title
                                            })
                                          }
                                        });

                                        resolve(promiseData);
                                      });
                                    } //response.ok
                                    else {
                                      //const statusCode = response.status;
                                      //let statusNum = response.status;
                                      //const statusMessage = response.statusMessage; //May not exist?
                                      response.json().then(data => {
                                        console.log(data);
                                        reject(data.error.message);
                                      })
                                      .catch (error => {
                                        //console.log("status: " + statusCode.toString() + " / " + statusNum.toString());
                                        reject(error.message);
                                        //Reset the promise cache in case of temp issue
                                        self.dataCache.views[item.list] = null;
                                      });
                                    }
                                  })
                                  .catch(error => {
                                    //.message: "Unexpected end of JSON input"
                                    //.stack: "SyntaxError: Unexpected end of JSON input\n    at e.json..."
                                    reject(error.message);
                                    //Reset the promise cache in case of temp issue
                                    self.dataCache.views[item.list] = null;
                                  });
                              });

                              //Store promise in cache and return
                              self.dataCache.views[item.list] = viewPromise;
                              return viewPromise;
                            }
                          })
                        )
                      }
                    },
                    {
                      id: "titleField",  
                      title: "Event Title",
                      required: true,
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately; not honoring deferredValidationTime on other fields
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: item.list,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null)
                                onUpdate(field.id, null);
                              else
                                onUpdate(field.id, option.key);
                            },
                            loadOptions: () => {
                              if (value == null)
                                onCustomFieldValidation(field.id, ''); //let the default "required" message show or type a custom message

                              //Look for an existing fields check
                              let fieldPromise = null as any;
                              if (self.dataCache.fields[item.list] == null) {
                                //Create a wrapper Promise to use later (needed to avoid an error calling response.json() multiple times)
                                fieldPromise = new Promise<[]>((resolve, reject) => {
                                  //Reference: https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ms428806(v=office.14)
                                    //OutputType == 2 per above for text-based Calculated columns
                                  
                                  //TODO: Add to $select: ,Choices
                                  //  so that if selected, show user modal asking if these should be added to list of Categories
                                  spHttpClient.get(item.siteUrl + "/_api/web/lists('" + item.list + "')/fields?$select=Id,InternalName,Title,ReadOnlyField,FieldTypeKind,TypeAsString&$filter=TypeAsString ne 'Computed' and Hidden eq false", SPHttpClient.configurations.v1) //was: ReadOnlyField eq false and 
                                    .then((response: SPHttpClientResponse) => {
                                      if (response.ok) {
                                        //TODO: Consider .text() here and then try/catch with JSON.parse
                                        response.json().then((data:any) => {
                                          resolve(data.value);
                                        });
                                      }
                                      else {
                                        //const statusCode = response.status;
                                        //let statusNum = response.status;
                                        //const statusMessage = response.statusMessage; //may not exist?
                                        response.json().then(data => {
                                          console.log(data);
                                          reject(data.error.message);
                                        })
                                        .catch (error => {
                                          //console.log("status: " + statusCode.toString() + " / " + statusNum.toString());
                                          reject(error.message);
                                          //Reset the promise cache in case of temp issue
                                          self.dataCache.fields[item.list] = null;
                                        });
                                      }
                                    })
                                    .catch(error => {
                                      //.message: "Unexpected end of JSON input"
                                      //.stack: "SyntaxError: Unexpected end of JSON input\n    at e.json..."
                                      reject(error.message);
                                      //Reset the promise cache in case of temp issue
                                      self.dataCache.fields[item.list] = null;
                                    });
                                });

                                //Store promise in cache
                                self.dataCache.fields[item.list] = fieldPromise;
                              }
                              else {
                                //Data promise already exists
                                fieldPromise = self.dataCache.fields[item.list];
                              }                              

                              //Build dropdown return promise
                              const returnPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                fieldPromise.then((fields:[]) => {
                                  let promiseData:IDropdownOption[] = [];
                                  //Add a blank option
                                  // promiseData.push({
                                  //   key: "", //blank
                                  //   text: ""
                                  // });
                                  //Add results to dropdown
                                  fields.forEach((field:any) => {
                                    //Only add applicable fields for the "Event Title"
                                    if (field.TypeAsString == "Calculated" || (field.ReadOnlyField == false && //Calculated is first because it's a ReadOnlyField
                                        (field.TypeAsString == "Text" || field.TypeAsString == "Choice" || field.TypeAsString == "Lookup" || 
                                           field.TypeAsString == "User")))
                                      promiseData.push({
                                        key: field.InternalName,
                                        text: field.Title,
                                        title: field.Title + (field.Title !== field.InternalName ? ` (${field.InternalName})` : '')
                                      });
                                  });

                                  resolve(promiseData);
                                })
                              });
                              return returnPromise;
                            }
                          })
                        )
                      }
                    },
                    {
                      id: "startDateField",  
                      title: "Start Date",
                      required: true,
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately; not honoring deferredValidationTime on other fields
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: item.list,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null)
                                onUpdate(field.id, null);
                              else
                                onUpdate(field.id, option.key);
                            },
                            loadOptions: () => {
                              if (value == null)
                                onCustomFieldValidation(field.id, ''); //let the default "required" message show
                              
                              //Get existing fields promise (created in "Event Title") field
                              const fieldPromise = self.dataCache.fields[item.list];
                              if (fieldPromise) {
                                //Build dropdown return promise
                                const returnPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                  fieldPromise.then((fields:[]) => {
                                    let promiseData:IDropdownOption[] = [];
                                    //Add a blank option
                                    // promiseData.push({
                                    //   key: "", //blank
                                    //   text: ""
                                    // });
                                    //Add results to dropdown
                                    fields.forEach((field:any) => {
                                      //Only add date fields
                                      if (field.TypeAsString == "DateTime")
                                        promiseData.push({
                                          key: field.InternalName,
                                          text: field.Title,
                                          title: field.Title + (field.Title !== field.InternalName ? ` (${field.InternalName})` : '')
                                        });
                                    });

                                    resolve(promiseData);
                                  })
                                });
                                return returnPromise;
                              }
                              else
                                return null;
                            }
                          })
                        )
                      }
                    },
                    {
                      id: "endDateField",  
                      title: "End Date",
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately; not honoring deferredValidationTime on other fields
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: item.list,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null)
                                onUpdate(field.id, null);
                              else
                                onUpdate(field.id, option.key);
                            },
                            loadOptions: () => {
                              //Get existing fields promise from startDate field
                              const fieldPromise = self.dataCache.fields[item.list];
                              if (fieldPromise) {
                                //Build dropdown return promise
                                const returnPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                  fieldPromise.then((fields:[]) => {
                                    let promiseData:IDropdownOption[] = [];
                                    //Add a blank option
                                    promiseData.push({
                                      key: "", //blank
                                      text: ""
                                    });
                                    //Add results to dropdown
                                    fields.forEach((field:any) => {
                                      //Only add date fields
                                      if (field.TypeAsString == "DateTime")
                                        promiseData.push({
                                          key: field.InternalName,
                                          text: field.Title,
                                          title: field.Title + (field.Title !== field.InternalName ? ` (${field.InternalName})` : '')
                                        });
                                    });

                                    resolve(promiseData);
                                  })
                                });
                                return returnPromise;
                              }
                              else
                                return null;
                            }
                          })
                        )
                      }
                    },
                    {
                      id: "category",  
                      title: "Category",
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately; not honoring deferredValidationTime on other fields
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: item.list,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null || (option != null && option.key == "")) {
                                //Clear related values (TODO: do this elsewhere)
                                // if (item.classField)
                                //   item.classField = null;
                                // if (item.className)
                                //   item.className = null;
                                onUpdate(field.id, null);
                              }
                              else
                                onUpdate(field.id, option.key); //other values set while list data is queried (in .tsx)
                            },
                            loadOptions: () => {
                              //Get existing fields promise from startDate field
                              const fieldPromise = self.dataCache.fields[item.list];
                              if (fieldPromise) {
                                //Build dropdown return promise
                                const returnPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                  fieldPromise.then((fields:[]) => {
                                    //Add a blank option
                                    let promiseData:IDropdownOption[] = [];
                                    promiseData.push({
                                      key: "", //blank
                                      text: ""
                                    });
                                    //Add fields header
                                    promiseData.push({
                                      key: "fieldsHeader",
                                      text: "Category Field",
                                      itemType: DropdownMenuItemType.Header
                                    });
                                    //Add results to dropdown
                                    fields.forEach((field:any) => {
                                      //Only add applicable fields
                                      // if (field.TypeAsString == "Calculated" || (field.ReadOnlyField == false && //Calculated is first because it's a ReadOnlyField
                                      //      (field.TypeAsString == "Text" || field.TypeAsString == "Choice" || field.TypeAsString == "Lookup" || 
                                      //         field.TypeAsString == "User")))
                                      if (field.TypeAsString == "Calculated" || (field.ReadOnlyField == false && //Calculated is first because it's a ReadOnlyField
                                        (field.TypeAsString == "Text" || field.TypeAsString == "Choice" || field.TypeAsString == "MultiChoice" || 
                                          field.TypeAsString == "Lookup" || field.TypeAsString == "LookupMulti" ||
                                          //OutcomeChoice is "Task Outcome in the classic UI"
                                          field.TypeAsString == "OutcomeChoice" || field.TypeAsString == "User" || field.TypeAsString == "UserMulti" ||
                                          field.TypeAsString == "TaxonomyFieldType" || field.TypeAsString == "TaxonomyFieldTypeMulti")))
                                        promiseData.push({
                                          key: "Field:" + field.InternalName,
                                          text: field.Title,
                                          title: field.Title + (field.Title !== field.InternalName ? ` (${field.InternalName})` : '')
                                        });
                                    });

                                    //Add static header
                                    promiseData.push({
                                      key: "staticHeader",
                                      text: "Static Category",
                                      itemType: DropdownMenuItemType.Header
                                    });
                                    //Add categories to dropdown
                                    if (this.properties.categories && this.properties.categories.length > 0)
                                      this.properties.categories.forEach((category:ICategoryItem) => {
                                        promiseData.push({
                                          key: "Static:" + category.uniqueId, 
                                          text: category.name
                                        });
                                      });
                                    else
                                      promiseData.push({
                                        key: "noStaticValues",
                                        text: "No categories created",
                                        disabled: true
                                      });

                                    resolve(promiseData);
                                  })
                                });
                                return returnPromise;
                              }
                              else
                                return null;
                            }
                          })
                        )
                      }
                    },
                    {
                      id: "group",  
                      title: "Row/Swimlane",
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately; not honoring deferredValidationTime on other fields
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: item.list,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null || (option != null && option.key == ""))
                                onUpdate(field.id, null);
                              else
                                onUpdate(field.id, option.key); //other values set while list data is queried (in .tsx)
                            },
                            loadOptions: () => {
                              //Get existing fields promise from startDate field
                              const fieldPromise = self.dataCache.fields[item.list];
                              if (fieldPromise) {
                                //Build dropdown return promise
                                const returnPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                  fieldPromise.then((fields:[]) => {
                                    //Add a blank option
                                    let promiseData:IDropdownOption[] = [];
                                    promiseData.push({
                                      key: "", //blank
                                      text: ""
                                    });
                                    //Add fields header
                                    promiseData.push({
                                      key: "fieldsHeader",
                                      text: "Row/Swimlane Field",
                                      itemType: DropdownMenuItemType.Header
                                    });
                                    //Add results to dropdown
                                    fields.forEach((field:any) => {
                                      //Only add applicable fields
                                      if (field.TypeAsString == "Calculated" || (field.ReadOnlyField == false && //Calculated is first because it's a ReadOnlyField
                                           (field.TypeAsString == "Text" || field.TypeAsString == "Choice" || field.TypeAsString == "MultiChoice" || 
                                            field.TypeAsString == "Lookup" || field.TypeAsString == "LookupMulti" ||
                                            //OutcomeChoice is "Task Outcome in the classic UI"
                                            field.TypeAsString == "OutcomeChoice" || field.TypeAsString == "User" || field.TypeAsString == "UserMulti" ||
                                            field.TypeAsString == "TaxonomyFieldType" || field.TypeAsString == "TaxonomyFieldTypeMulti")))
                                        promiseData.push({
                                          key: "Field:" + field.InternalName,
                                          text: field.Title,
                                          title: field.Title + (field.Title !== field.InternalName ? ` (${field.InternalName})` : '')
                                        });
                                    });

                                    //Add static header
                                    promiseData.push({
                                      key: "staticHeader",
                                      text: "Static Row/Swimlane",
                                      itemType: DropdownMenuItemType.Header
                                    });
                                    //Add rows/swimlanes to dropdown
                                    if (this.properties.groups && this.properties.groups.length > 0)
                                      this.properties.groups.forEach((group: IGroupItem, index) => {
                                        //When sortIdx doesn't exist, it's a new row
                                        if (item.sortIdx == null && index == 0 && item.group == null) { //When null this means it's a new row, so default select the first group so items will render somewhere
                                          promiseData.push({
                                            //key:group.uniqueId,
                                            key: "Static:" + group.uniqueId, 
                                            text:group.name,
                                            selected:true
                                          });
                                          item.group = "Static:" + group.uniqueId; //needed to actually have a value set in case the user doesn't change it
                                        }
                                        else
                                          promiseData.push({
                                            key:"Static:" + group.uniqueId,
                                            text:group.name
                                          });
                                      });
                                    else
                                      promiseData.push({
                                        key: "noStaticValues",
                                        text: "No rows created",
                                        disabled: true
                                      });

                                    resolve(promiseData);
                                  })
                                });
                                return returnPromise;
                              }
                              else
                                return null;
                            }
                          })
                        )
                      }
                    },
                    /*{
                      id: "groupId",
                      title: "Row/Swimlane",
                      isVisible: (field:ICustomCollectionField, items:IListItem[]):boolean => {
                        if (this.properties.groups && this.properties.groups.length > 0)
                          return true;
                        else
                          return false;
                      },
                      disable: (item:IListItem):boolean => {
                        return (item.list ? false : true);
                      },
                      placeholder: " ", //need a space because blank just shows the title
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      //NOTE: Only fired when *initialy* rendered, not after other fields are changed
                      options: (fieldId: string, item: IListItem) => {
                        let options: ICustomDropdownOption[] = [{key: "", text: ""}]; //adding a blank entry
                        if (this.properties.groups)
                          this.properties.groups.forEach((group: IGroupItem, index) => {
                            if (index == 0 && item.groupId == null) { //for new PropFieldCollection rows, select the first group
                              options.push({key:group.uniqueId, text:group.name, selected:true });
                              item.groupId = group.uniqueId; //needed to actually have a value set in case the user doesn't change it
                            }
                            else
                              options.push({key:group.uniqueId, text:group.name });
                          });
                        return options;
                      }
                    },*/
                    // { //Moved to be a prop within Advanced Configs
                    //   id: "visible",
                    //   title: "Enabled",
                    //   type: CustomCollectionFieldType.boolean,
                    //   defaultValue: true,
                    // },
                    {
                      id: "configs",  
                      title: "Advanced Configs",
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item:IListItem, itemId, onCustomFieldValidation) => {  
                        //Provide a default value to show in the editor
                        if (value == null || value == "")
                          value = "{\r\n  \"visible\": true\r\n}";

                        return (
                          React.createElement(MonacoPanelEditor, {
                            key: itemId,
                            disabled: (item.list ? false : true),
                            buttonText: "Advanced",
                            headerText: 'Advanced JSON attribute editor for SharePoint List configuration',
                            value: value,
                            language: "json",
                            onValueChanged: (newValue: string) => {
                              //Proceed saving the data only if it's valid JSON
                              try {
                                JSON.parse(newValue); //exception if not
                                onUpdate(field.id, newValue); //save the value
                              }
                              catch (e) {
                                //Nothing needed
                              }
                            }
                          })
                        )
                      }
                    }
                  ]
                }),
                PropertyFieldCollectionData("calsAndPlans", {
                  key: "calsAndPlans",
                  value: this.properties.calsAndPlans,
                  tableClassName: "calendarsFCDTable",
                  label: "Outlook Calendars", //Header/label above the button
                  manageBtnLabel: "Add/Edit Calendars",
                  panelProps: {
                    type: PanelType.smallFluid
                  },
                  panelHeader: "Configure Outlook Calendars",
                  panelDescription: "Specify the desired calendars and the category to be assigned.",
                  saveBtnLabel: "Save & close",
                  saveAndAddBtnLabel: "Add/Save & close",
                  //enableSorting: true, //not necessary here as the list order doesn't matter
                  fields: [
                    {
                      id: "persona",
                      title: "User/Org Box/Group",
                      required: true,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item:ICalendarItem, itemId, onCustomFieldValidation) => {  
                        return (  
                          React.createElement(DirectoryPicker, {
                            //key: itemId,
                            //uniqueId: item.uniqueId, //for testing
                            //sortIdx: item.sortIdx, //for testing
                            disabled: false,
                            selectedPersonas: value,
                            initialSuggestions: this.dataCache.memberOf,
                            graphClient: this._graphClient,
                            getGraphScopes: this.getGraphScopes.bind(this),
                            onChange: (items:IPersonaProps[]) => {
                              if (items.length == 0) {
                                //neither helped
                                onCustomFieldValidation(field.id, ''); //"" empty doesn't prevent
                                onUpdate(field.id, null);
                              }
                              else {
                                //Create a "trimmer" object
                                let item = [{
                                  key: items[0].key,
                                  text: items[0].text,
                                  mail: items[0].mail,
                                  personaType: items[0].personaType
                                } as IPersonaProps];
                                onCustomFieldValidation(field.id, '');
                                onUpdate(field.id, item);
                              }
                            }
                          })
                        )
                      }
                    },
                    {
                      id: "resource",  
                      title: "Calendar",
                      required: true,
                      type: CustomCollectionFieldType.custom,
                      //NOTE: Fired immediately; not honoring deferredValidationTime on other fields
                      onCustomRender: (field, value, onUpdate, item:ICalendarItem, itemId, onCustomFieldValidation) => {
                        return (  
                          React.createElement(AsyncDropdown, {
                            label: undefined,
                            selectedKey: value,
                            disabled: false,
                            stateKey: (item.persona && item.persona[0] ? item.persona[0].key : null),
                            forceNoDelay: true,
                            onChange: (event:Event, option: IDropdownOption) => {
                              if (option == null || (option != null && option.key == "")) {
                                onUpdate(field.id, null);
                                //onCustomFieldValidation(field.id, "Calendar must be selected"); //not needed as default seems to automatically be shown
                              }
                              else {
                                onUpdate(field.id, option.key);

                                //For user calendars, need to make sure current user has at least "view all details" permission or access denied error will be thrown getting events
                                const persona = item.persona[0];
                                if (persona.personaType == "user")
                                  this._graphClient.then((client:MSGraphClientV3): void => {
                                    const now = new Date();
                                    let later = new Date();
                                    later.setDate(later.getDate() + 1);
                                    //Just try to get any event to see if access denied is returned
                                    client.api("/users/" + persona.key + "/calendars/" + option.key + "/calendarView")
                                    .query(`startDateTime=${now.toISOString()}&endDateTime=${later.toISOString()}`)
                                    .select("id,subject").top(1)
                                    .get((error:GraphError, response:any, rawResponse?:any) => {
                                      if (error && error.code == "ErrorAccessDenied") {
                                        onCustomFieldValidation(field.id, 'You need at least the "view all details" permission to this calendar');
                                      }
                                    })
                                    .catch(reason => { //reason is undefined
                                      //Just catch to prevent "Uncaught (in promise)" console error
                                    })
                                  });
                              }
                            },
                            loadOptions: () => {
                              const personaId = (item.persona[0] ? item.persona[0].key : null);
                              const persona = item.persona[0];
                              //onCustomFieldValidation(field.id, "At load options");
                              
                              //Look for an existing promise
                              if (self.dataCache.calendars[personaId]) {
                                /*const test1 = self.dataCache.calendars[personaId];
                                (test1 as Promise<IDropdownOption[]>).catch(reason => {
                                  console.warn(reason);
                                  onCustomFieldValidation(field.id, "in reused promise catch");
                                });*/
                                return self.dataCache.calendars[personaId];
                              }

                              //For groups just return a "default" calendar (since there cannot be other "created" calendars like with users)
                              if (persona.personaType == "group") {
                                //Just "return" this simple promise
                                return new Promise<IDropdownOption[]>((resolve, reject) => {
                                  let promiseData:IDropdownOption[] = [];
                                  promiseData.push({
                                    key: "calendar:default",
                                    text: "Calendar"
                                  });
                                  resolve(promiseData);
                                });

                                //TODO: Query for Planner plans
                                //key: plan:{id}
                              }

                              //For users, query to see which calendars are available/shared
                              const calendarsPromise = new Promise<IDropdownOption[]>((resolve, reject) => {
                                let errorMsg = null as string;
                                //Get data
                                this._graphClient.then((client:MSGraphClientV3): void => {
                                  client.api("/users/" + persona.key + "/calendars").select("id,name,isDefaultCalendar,owner")
                                  .get((error:GraphError, response:any, rawResponse?:any) => {
                                    if (error) {
                                      //When Calendars.Read[.*] Graph scope is not approved...
                                      //code: "ErrorAccessDenied"
                                      //message: "Access is denied. Check credentials and try again."

                                      //Reset the promise cache in case of temp issue
                                      self.dataCache.calendars[personaId] = null;
                                      
                                      //Provide a clearer error message in case of no permissions to calendar
                                      if (error.message == "The specified object was not found in the store.") {
                                        //Store error for use inside catch()
                                        errorMsg = "No permissions to load calendars for selected user";
                                        reject("No permissions to load calendars for selected user");
                                        //This works here if no rejection
                                        //onCustomFieldValidation(field.id, "message");
                                      }
                                      else {
                                        errorMsg = error.message;
                                        reject(error.message);
                                      }
                                    }
                                    else {
                                      let promiseData:IDropdownOption[] = [];
                                      // promiseData.push({
                                      //   key: "", //blank
                                      //   text: ""
                                      // });

                                      const calendars:MicrosoftGraph.Calendar[] = response.value;
                                      calendars.forEach(calendar => {
                                        //Exclude other users' calendars shared with the selected user
                                        if (calendar.owner.address != persona.mail)
                                          return;

                                        //Ignore these known calendars also in case the person selects themself
                                        if (calendar.name != "United States holidays" && calendar.name != "Birthdays")
                                          promiseData.push({
                                            key: "calendar:" + calendar.id,
                                            text: calendar.name
                                          });
                                      });
                                      resolve(promiseData);
                                      onCustomFieldValidation(field.id, ''); //Show default msg
                                    }
                                  })
                                  .catch(reason => { //even with no reject above, this is undefined
                                    onCustomFieldValidation(field.id, errorMsg);
                                  })

                                  /* None of these worked
                                  .get().then(response => {
                                    let promiseData:IDropdownOption[] = [];
                                    promiseData.push({
                                      key: "", //blank
                                      text: ""
                                    });

                                    const calendars:MicrosoftGraph.Calendar[] = response.value;
                                    calendars.forEach(calendar => {
                                      //Exclude other users' calendars shared with the selected user
                                      if (calendar.owner.address != persona.mail)
                                        return;

                                      //Ignore these known calendars also in case the person selects themself
                                      if (calendar.name != "United States holidays" && calendar.name != "Birthdays")
                                        promiseData.push({
                                          key: calendar.id,
                                          text: calendar.name
                                        });
                                    });
                                    resolve(promiseData);
                                  //Handle any errors
                                  })//, failureReason => {
                                  .catch((error:GraphError) => {
                                    // If only given "Can view titles and locations" (not Can view all details+),
                                    //   when querying the calendar this error is produced:
                                    // code: "ErrorAccessDenied"
                                    // message: "Access is denied. Check credentials and try again."

                                    //Reset the promise cache in case of temp issue
                                    self.dataCache.calendars[personaId] = null;
                                          
                                    //Provide a clearer error message in case of no permissions to calendar
                                    if (error.message == "The specified object was not found in the store.") {
                                      //reject("something");
                                      onCustomFieldValidation(field.id, "No permissions to load calendars for selected user");
                                    }
                                    else {
                                      onCustomFieldValidation(field.id, error.message);
                                      reject();
                                    }
                                  })*/
                                });
                              });

                              //Providing no value
                              // calendarsPromise.catch(reason => {
                              //   console.log(reason); //has access to reject msg
                              //   onCustomFieldValidation(field.id, "bottom catch"); //not shown (because Uncaught (in promise) error or get catch message shown instead)
                              // });

                              //Store promise in cache and return
                              self.dataCache.calendars[personaId] = calendarsPromise;
                              return calendarsPromise;
                            }
                          })
                        )
                      }
                    },
                    //Custom filter query for calendar items
                    {
                      id: "filter",
                      title: "Filter Query",
                      /*isVisible: (field:ICustomCollectionField, items:ICalendarItem[]):boolean => {
                        if (this.properties.groups && this.properties.groups.length > 0)
                          return true;
                        else
                          return false;
                      },*/
                      disable: (item:ICalendarItem):boolean => {
                        return (item.persona == null || item.persona.length == 0 || item.resource == null ? true : false);
                      },
                      placeholder: " ", //need a space because blank just shows the title
                      type: CustomCollectionFieldType.string,
                      required: false,
                      deferredValidationTime: 1000,
                      //Oddly named: This is really the "perform field validation" function
                      onGetErrorMessage: (value: string, index: number, item: ICalendarItem) => {
                        //NOTE: "this" is just the field object
                        //Fired after deferredValidationTime

                        //Handle blank and cleared-out values
                        if (value == null || value.trim() == '')
                          return ''; //no validation error; '' lets default checks happen

                        //Look for existing check and return it's promise
                        if (self.dataCache.calFilterQuery[value])
                          return self.dataCache.calFilterQuery[value]

                        //Try a calendar Graph call to test if query is valid
                        const promise = new Promise<string>((resolve, reject) => {
                          this._graphClient.then((client:MSGraphClientV3): void => {
                            //Build API URL based on user or group calendar
                            let apiURL = "";
                            let resourceId = item.resource.split(":")[1]; //format "calendar:id"
                            if (item.persona[0].personaType == "user")
                              apiURL = "/users/" + item.persona[0].mail + "/calendars/" + resourceId + "/calendarView";
                            else //assumed to be a group
                              apiURL = "/groups/" + item.persona[0].key + "/calendarView";

                            const now = new Date();
                            let later = new Date();
                            later.setDate(later.getDate() + 1);
                            
                            //Get sample calender view events just to test the input query
                            //return client.api(apiURL).query(`startDateTime=${this.getMinDate().toISOString()}&endDateTime=${this.getMaxDate().toISOString()}`)
                            client.api(apiURL).query(`startDateTime=${now.toISOString()}&endDateTime=${later.toISOString()}`)
                            .select("id,subject").top(1)
                            .filter(value.trim())
                            .get((error:GraphError, response:any, rawResponse?:any) => {
                              if (error) {
                                resolve(error.message);
                              }
                              else {
                                resolve(''); //no validation error
                              }
                            })
                            .catch(reason => { //reason is undefined
                              //Just catch to prevent "Uncaught (in promise)" console error
                            });
                          });
                        });

                        //Store the promise in cache and return
                        self.dataCache.calFilterQuery[value] = promise;
                        return promise;
                      }
                    },
                    {
                      id: "category",
                      title: "Category",
                      disable: (item:ICalendarItem):boolean => {
                        return (item.persona == null || item.persona.length == 0 ? true : false);
                      },
                      placeholder: " ", //need a space because blank just shows the title
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      //NOTE: Only fired when *initialy* rendered, not after other fields are changed
                      options: (fieldId: string, item: ICalendarItem) => {
                        let options: ICustomDropdownOption[] = [
                          //Add blank entry
                          {key: "", text: ""},
                          //Add fields header
                          {
                            key: "fieldsHeader",
                            text: "Category Field",
                            itemType: DropdownMenuItemType.Header
                          },
                          //Add fields from Outlook events
                          { key: "Field:categories", text: "Categories" },
                          { key: "Field:showAs", text: "Show As" },
                          { key: "Field:charmIcon", text: "Charm/Icon" },

                          //Add static header
                          {
                            key: "staticHeader",
                            text: "Static Category",
                            itemType: DropdownMenuItemType.Header
                          }
                        ];
                        
                        //Add categories to dropdown
                        if (this.properties.categories && this.properties.categories.length > 0)
                          this.properties.categories.forEach((category:ICategoryItem) => {
                            options.push({
                              key: "Static:" + category.uniqueId, 
                              text: category.name
                            });
                          });
                        else
                          options.push({
                            key: "noStaticValues",
                            text: "No categories created",
                            disabled: true
                          });

                        return options;
                      }
                    },
                    {
                      id: "group",
                      title: "Row/Swimlane",
                      isVisible: (field:ICustomCollectionField, items:ICalendarItem[]):boolean => {
                        if (this.properties.groups && this.properties.groups.length > 0)
                          return true;
                        else
                          return false;
                      },
                      disable: (item:ICalendarItem):boolean => {
                        return (item.persona == null || item.persona.length == 0 ? true : false);
                      },
                      placeholder: " ", //need a space because blank just shows the title
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      //NOTE: Only fired when *initialy* rendered, not after other fields are changed
                      options: (fieldId: string, item: IListItem) => {
                        let options: ICustomDropdownOption[] = [
                          //Add blank entry
                          {key: "", text: ""},
                          //Add fields header
                          {
                            key: "fieldsHeader",
                            text: "Row/Swimlane Field",
                            itemType: DropdownMenuItemType.Header
                          },
                          //Add fields from Outlook events
                          { key: "Field:categories", text: "Categories" },
                          { key: "Field:showAs", text: "Show As" },
                          { key: "Field:charmIcon", text: "Charm/Icon" },
                          
                          //Add static header
                          {
                            key: "staticHeader",
                            text: "Static Row/Swimlane",
                            itemType: DropdownMenuItemType.Header
                          }
                        ];

                        //Add rows/swimlanes to dropdown
                        if (this.properties.groups && this.properties.groups.length > 0)
                          this.properties.groups.forEach((group: IGroupItem, index) => {
                            if (index == 0 && item.group == null) { //When null this means it's a new row, so default select the first group so items will render somewhere
                              options.push({
                                //key:group.uniqueId,
                                key: "Static:" + group.uniqueId, 
                                text:group.name,
                                selected:true
                              });
                              item.group = "Static:" + group.uniqueId; //needed to actually have a value set in case the user doesn't change it
                            }
                            else
                            options.push({
                                key:"Static:" + group.uniqueId,
                                text:group.name
                              });
                          });
                        else
                        options.push({
                            key: "noStaticValues",
                            text: "No rows created",
                            disabled: true
                          });

                        return options;
                      }
                    },
                    {
                      id: "configs",  
                      title: "Advanced Configs",
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item:ICalendarItem, itemId, onCustomFieldValidation) => {  
                        //Provide a default value to show in the editor
                        if (value == null || value == "")
                          value = "{\r\n  \"visible\": true\r\n}";

                        return (
                          React.createElement(MonacoPanelEditor, {
                            key: itemId,
                            disabled: (item.persona == null || item.persona.length == 0 ? true : false),
                            buttonText: "Advanced",
                            headerText: 'Advanced JSON attribute editor for Outlook Calendar configuration',
                            value: value,
                            language: "json",
                            onValueChanged: (newValue: string) => {
                              //Proceed saving the data only if it's valid JSON
                              try {
                                JSON.parse(newValue); //exception if not
                                onUpdate(field.id, newValue); //save the value
                              }
                              catch (e) {
                                //Nothing needed
                              }
                            }
                          })
                        )
                      }
                    }
                  ]
                })
              ]
            },
            {
              groupName: "Visualization Settings",
              groupFields: [
                PropertyPaneTextField('holidayCategories', { //TODO: allow multiple, delimited values?
                  label: "Holiday category",
                  //value: "Holiday", //set in WP manifest
                  description: "Category that will render as a vertical background bar"
                  //deferredValidationTime: 1000 //only applies to validation; still *immediately* fires onPropertyPaneFieldChanged
                }),
                PropertyPaneToggle('fillFullWidth', {
                  key: "fillFullWidth",
                  label: "Fill the full width of available page size",
                  checked: false
                }),
                PropertyPaneToggle('calcMaxHeight', {
                  key: "calcMaxHeight",
                  label: "Set a fixed height to available page size",
                  checked: false
                }),
                PropertyPaneToggle('singleDayAsPoint', {
                  key: "singleDayAsPoint",
                  label: "Single day events show as a point/dot"
                  //checked: true //not needed if set in manifest.json file
                }),
                PropertyPaneToggle('overflowTextVisible', {
                  key: "overflowTextVisible",
                  label: "Allow event text to flow outside of margin"
                }),
                PropertyPaneToggle('hideItemBoxBorder', {
                  key: "hideItemBoxBorder",
                  label: "Hide event box borders (only show a line)"
                })
                //Not wanting to offer these yet
                // PropertyPaneToggle('hideSocialBar', {
                //   label: "Hide social/comments area at page bottom",
                //   checked: false
                // })
                /*PropertyPaneToggle('getDatesAsUtc', {
                  label: "Convert list/event data to Zulu/UTC time",
                  checked: false
                })*/
              ]
            },
            {
              groupName: "Date Settings",
              groupFields: [
                PropertyFieldMessage("", {
                  key: "dateSettingsMsg",
                  text: "Note that these settings won't cause an immediate refresh of the timeline",
                  messageType: MessageBarType.info,
                  isVisible: true
                }),
                PropertyFieldNumber('initialStartDays', {
                  key: 'initialStartDays',
                  label: 'Days in past to render the initial view',
                  description: 'Beginning day shown for the initial timeline load',
                  value: this.properties.initialStartDays,
                  minValue: 0,
                  maxValue: 365
                }),
                PropertyFieldNumber('initialEndDays', {
                  key: 'initialEndDays',
                  label: 'Days in future to render the initial view',
                  description: 'Ending day shown for the initial timeline load',
                  value: this.properties.initialEndDays,
                  minValue: 1
                }),
                PropertyFieldNumber('minDays', {
                  key: 'minDays',
                  label: 'Days in past able to scroll back in time',
                  description: 'Earliest day you can scroll to the left',
                  value: this.properties.minDays,
                  minValue: 0,
                  maxValue: 365,
                  precision: 0 //only whole number stored
                }),
                PropertyFieldNumber('maxDays', {
                  key: 'maxDays',
                  label: 'Days in future able to scroll ahead in time',
                  description: 'Latest day you can scroll to the right',
                  value: this.properties.maxDays,
                  minValue: 1,
                  precision: 0
                })
              ]
            }
          ]
        },
        {
          /*header: {
            description: "page header"
          },*/
          groups: [
            {
              groupName: "Advanced Settings",
              groupFields: [
                PropertyPaneLabel('visJsonProperties', {
                  text: "Edit Timeline visualization properties"
                }),
                PropertyPaneWebPartInformation({ //was adding to the div: style="font-size:.9em;"
                  description: `<div>Refer to the <a href="https://visjs.github.io/vis-timeline/docs/timeline/#Configuration_Options" target="_blank">Timeline configuration options</a> page for available options</div>`,
                  key: 'visInstructions'
                }),
                PropertyFieldMonacoEditor('visJsonProperties', {
                  key: 'visJsonProperties',
                  value: this.properties.visJsonProperties,
                  showMiniMap: false,
                  showLineNumbers: true,
                  onChange: (newValue:string) => {
                    //Fired even when *cancel* button is clicked (but newValue is not what user typed/changed)
                    //Data is already saved at this point (no need to manually save it here)
                    //newValue is what the user typed, even if the prop was overwritten in onPropertyPaneFieldChanged
                    //Function must exist to prevent error clicking Cancel button
                  },
                  language: "json", //css, html, json, typescript
                  theme: "vs-dark"
                }),
                PropertyPaneLabel('', { //Just using for spacing
                  text: ""
                }),
                PropertyPaneLabel('tooltipLabel', {
                  text: "Customize the hover-over tooltip"
                }),
                PropertyPaneWebPartInformation({
                  description: `<div>Refer to the <a href="https://handlebarsjs.com/guide/#language-features" target="_blank">Handlebars Language Guide</a> page</div>`,
                  key: 'ttInstructions'
                }),
                PropertyFieldMonacoEditor('tooltipEditor', {
                  key: 'tooltipEditor',
                  value: this.properties.tooltipEditor,
                  showMiniMap: false,
                  showLineNumbers: true,
                  onChange: (newValue:string) => {
                    //Function must exist to prevent error clicking Cancel button
                  },
                  language: "html", //css, html, json, typescript
                  theme: "vs-dark"
                }),
                PropertyPaneLabel('', { //Just using for spacing
                  text: ""
                }),
                PropertyPaneLabel('cssLabel', {
                  text: "Add custom CSS overrides"
                }),
                PropertyPaneWebPartInformation({
                  description: `<div>Use this to add custom CSS to the page</div>`,
                  key: 'cssInstructions'
                }),
                PropertyFieldMonacoEditor('cssOverrides', {
                  key: 'cssOverrides',
                  value: this.properties.cssOverrides,
                  showMiniMap: false,
                  showLineNumbers: true,
                  onChange: (newValue:string) => {
                    //Function must exist to prevent error clicking Cancel button
                  },
                  language: "css", //css, html, json, typescript
                  theme: "vs-dark"
                })
              ]
            }
          ]
        },
        {
          /*header: {
            description: "About page header"
          },*/
          groups: [
            {
              groupName: "About",
              groupFields: [
                PropertyPaneWebPartInformation({
                  //No margin-top style needed
                  description: `<div style="margin-bottom:5px"><b>Reference & Support</b></div>
                    <div style="margin-bottom:5px">Use the following links to access documentation and support as well as to report any issues or to submit an idea for a new feature.</div>`,
                  key: 'supportInfo'
                }),
                PropertyPaneLink('',{
                  target: '_blank',
                  href: "https://www.milsuite.mil/book/groups/m365-support/projects/timeline-calendar/",
                  text: "milBook Group/Project (DoD CAC-login)"
                }),
                PropertyPaneLink('',{
                  target: '_blank',
                  href: "https://github.com/spsprinkles/timeline-calendar/",
                  text: "GitHub Repository (public access)"
                }),
                // PropertyPaneMarkdownContent({
                //   markdown: webpartMD,
                //   key: "webpartInfo"
                // }),
                PropertyPaneWebPartInformation({
                  description: `<div style="margin-top:20px;margin-bottom:5px;"><b>Web Part Version</b></div>
                    <div>${this && this.manifest.version ? this.manifest.version : '*Unknown*'}</div>
                    <div style="margin-top:20px;margin-bottom:5px;"><b>Web Part Instance ID</b></div>
                    <div>${this.instanceId}</div>`,
                  key: 'wpInfo'
                }),
                PropertyPaneWebPartInformation({
                  description: `<div style="margin-top:20px;margin-bottom:5px;"><b>Author</b></div>
                    <div>Michael Vasiloff <a href="https://www.linkedin.com/in/michaelvasiloff" target="_blank">[LinkedIn]</a> <a href="https://github.com/mikevasiloff" target="_blank">[GitHub]</a> <a href="https://www.milsuite.mil/book/people/michael.d.vasiloff" target="_blank">[milBook]</a></div>`,
                  key: 'authors'
                }),
                PropertyPaneWebPartInformation({
                  description: `<div style="margin-top:20px;margin-bottom:5px;"><b>Graph API Scopes (Tenant Approved)</b></div>
                    <div>${this.getGraphScopes(true).join(", ")}</div>`,
                  key: 'graphScopes'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
