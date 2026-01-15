import * as React from 'react';
import * as ReactDom from 'react-dom';
//import styles from './TimelineCalendar.module.scss';
import { ITimelineCalendarProps } from './ITimelineCalendarProps';
//import { escape } from '@microsoft/sp-lodash-subset';
//import window as any;

import 'vis-timeline/dist/vis-timeline-graph2d.min.css';
import './VisStyleOverrides';
import { DataSet } from 'vis-data';
import { Timeline } from 'vis-timeline'; //, TimelineOptions
//import { TagItemSuggestion } from 'office-ui-fabric-react';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import { IFrameDialog, IFrameDialogProps } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { DialogType } from 'office-ui-fabric-react/lib/Dialog';
//import { SPComponentLoader } from '@microsoft/sp-loader';
import { ICalendarConfigs, ICalendarItem, ICategoryItem, IGroupItem, IListConfigs, IListItem } from './IConfigurationItems';
import * as Handlebars from 'handlebars';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client'; //ResponseType
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
//import { DescriptionFieldLabel } from 'TimelineCalendarWebPartStrings';
//import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
//import { DefaultButton } from '@fluentui/react/lib/Button'; //PrimaryButton
//import { TeachingBubbleContentBase } from 'office-ui-fabric-react';
import { filterXSS, whiteList } from 'xss';

//declare const window: any; //temp TODO

class IdSvc {
  private static _id = 0;
  public static getNext(): number {
    this._id++;
    return this._id;
  }
  //private constructor() {}
}

interface IItemDateInfo {
  eventStartDate: Date
  eventEndDate?: Date
}

export default class TimelineCalendar extends React.Component<ITimelineCalendarProps, {}> {
  private _timeline: Timeline;
  private _dsItems: any;
  private _dsGroups: any;
  private _isLoadingEvents: boolean = false;

  /**
   * Called when component is mounted (only on the *initial* loading of the web part)
   */
  public async componentDidMount(): Promise<void> {
    //const { data, calendars } = this.props;
    this._isLoadingEvents = false; //Ensure flag is reset on mount
    this.initialBuildTimeline();

    //Add a helper for when user mistypes a helper name so that an exception is not thrown
    Handlebars.registerHelper('helperMissing', function( /* dynamic arguments */) {
      if (arguments.length === 1) {
        //This is actually just a field property *without* a value; no "helper" was specified
        return "";
      }
      else {
        //The handler name used doesn't exist
        const options = arguments[arguments.length-1];
        //const args = Array.prototype.slice.call(arguments, 0, arguments.length-1);
        return new Handlebars.SafeString('Missing handler: "' + options.name + '"');
      }
    });
    
    Handlebars.registerHelper("limit", strText => {
      if (strText == null)
        return "";
      else {
        const divWrapper = document.createElement("div");
        //Put description html into element so we can use querySelectorAll
        divWrapper.innerHTML = strText;
        divWrapper.querySelectorAll("*").forEach((elem: HTMLElement) => {
          if (elem.textContent && elem.textContent.length > 500)
            elem.textContent = elem.textContent.substring(0, 500) + "...";
        });
        return divWrapper.innerHTML;
      }
    });

    //TODO: Look at adding dateFormat also: https://github.com/tcort/handlebars-dateformat/blob/master/index.js
    //                                    https://docs.celigo.com/hc/en-us/articles/360045564992-Handlebar-expressions-for-date-and-time-format-codes
    Handlebars.registerHelper("date", date => {
      //Check for date fields
      const dateValue = new Date(date); //in case it's a string (works fine if already a Date)
      if (isNaN(dateValue.getTime()))
        return date; //original [string] value
      else {
        //Build a string in the format: "2023-10-10 08:00:00"
        let dateStr = dateValue.getFullYear().toString() + "-";
        if (dateValue.getMonth() < 9) //Month is 0-based index...
          dateStr += "0";
        dateStr += (dateValue.getMonth()+1).toString() + "-"; //...need to +1 to get actual month
        if (dateValue.getDate() < 10)
          dateStr += "0";
        dateStr += dateValue.getDate().toString() + " ";
        //Add time element
        if (dateValue.getHours() < 10)
          dateStr += "0";
        dateStr += dateValue.getHours().toString() + ":";
        if (dateValue.getMinutes() < 10)
          dateStr += "0";
        dateStr += dateValue.getMinutes().toString() + ":";
        if (dateValue.getSeconds() < 10)
          dateStr += "0";
        dateStr += dateValue.getSeconds().toString();
        return dateStr;
      }
    });

    Handlebars.registerHelper("yesNo", strText => {
      if (strText == null)
        return "";
      else {
        if (strText === "1")
          return "Yes";
        else
          return "No";
      }
    });
  }

  /**
   * Called when component is updated (for example, when properties are changed)
   * Fired *immediately* when props are changed, regardless of their deferredValidationTime value
   */
  public componentDidUpdate(prevProps: Readonly<ITimelineCalendarProps>, prevState: Readonly<{}>, snapshot?: any): void {
    //this.context.domElement is undefined here
    const self = this;
    let reloadEvents = true;
    //Check for specific property changes not requiring event reload
    // if (prevProps.categories != this.props.categories)
    //   reloadEvents = false;
    if (prevProps.groups != this.props.groups) {
      reloadEvents = false;
      //Check if there were no groups but now new ones were added; need to tag existing events to a group or they won't show
      if ((prevProps.groups == null || prevProps.groups.length === 0) && (this.props.groups && this.props.groups.length > 0)) { // && this.props.lists
        const groupId = (this.props.groups[0] as IGroupItem).uniqueId;
        const itemEvents = this._dsItems.get({ //get all events (except for "weekends")
          filter: function (item:any) {
            if (item.className !== "weekend") {
              item.group = groupId;
              return true;
            }
            return false;
          }
        });
        this._dsItems.update(itemEvents);
      }
    }
    else if (prevProps.holidayCategories != this.props.holidayCategories) {
      reloadEvents = false;
      const groupId = (self.props.groups == null ? null : (self.props.groups[0] as IGroupItem).uniqueId);
      //Update existing, applicable events type to "background"
      const itemEvents = this._dsItems.get({
        filter: function (item:any) {
          if (prevProps.holidayCategories != null && prevProps.holidayCategories !== "" && item.className === self.props.ensureValidClassName(prevProps.holidayCategories)) {
            item.type = "range"; //assume it should be reverted to range (vs. point)
            item.group = groupId;
            return true;
          }
          else if (item.className === self.props.ensureValidClassName(self.props.holidayCategories)) {
            item.type = "background"; //change to background
            item.group = null; //ensure holiday covers all groups
            return true;
          }
          else
            return false;
        }
      });
      this._dsItems.update(itemEvents);
      return; //don't run any of the below
    }
    else if (prevProps.initialStartDays != this.props.initialStartDays)
      return;
    else if (prevProps.initialEndDays != this.props.initialEndDays)
      return;
    else if (prevProps.minDays != this.props.minDays) {
      this._timeline.setOptions({min: this.getMinDate()});
      return;
    }
    else if (prevProps.maxDays != this.props.maxDays) {
      this._timeline.setOptions({max: this.getMaxDate()});
      return;
    }
    else if (prevProps.calcMaxHeight != this.props.calcMaxHeight) {
      if (this.props.calcMaxHeight)
        this._timeline.setOptions({maxHeight: this.calcMaxHeight()});
      else
        this._timeline.setOptions({maxHeight: ""}); //same as initially not setting this prop at all
      return;
    }
    else if (prevProps.fillFullWidth != this.props.fillFullWidth) {
      this.renderDynamicStyles();
      return;
    }
    else if (prevProps.singleDayAsPoint != this.props.singleDayAsPoint) {
      //Update events
      const itemEvents = this._dsItems.get({
        filter: (item:any) => {
          if (item.type != "background") { //ignore weekend and special holiday events
            //Was checking (item.start.toISOString().substring(0, 10) == item.end.toISOString().substring(0, 10)), but the ISO time zone threw off some events
            if (this.props.singleDayAsPoint && (item.end == null || item.start.toLocaleDateString() === item.end.toLocaleDateString())) //localDate == "11/27/2023"
              item.type = "point";
            else
              item.type = "range";

            return true;
          }
          else
            return false;
        }
      });
      this._dsItems.update(itemEvents);
      return; //don't run any of the below
    }
    else if (prevProps.hideItemBoxBorder != this.props.hideItemBoxBorder) {
      this.renderDynamicStyles();
      return;
    }
    else if (prevProps.overflowTextVisible != this.props.overflowTextVisible) {
      this.renderDynamicStyles();
      if (this.props.overflowTextVisible) {
        this._timeline.redraw(); //This alone doesn't fix issue with overlapping event text
        //Need to "move" the timeline window to force a redraw
        const tcWin = this._timeline.getWindow();
        tcWin.end.setSeconds(tcWin.end.getSeconds() + 1);
        this._timeline.setWindow(tcWin.start, tcWin.end);
      }
      return;
    }
    // else if (prevProps.hideSocialBar != this.props.hideSocialBar) {
    //   this.renderDynamicStyles();
    //   return;
    // }
    else if (prevProps.tooltipEditor != this.props.tooltipEditor) {
      //No need to overwrite this._timeline.setOptions because tooltip.template is already a dynamic function
      //Get original "field keys"
      const prevFieldKeys = this.getFieldKeys(prevProps.tooltipEditor);
      const newFieldKeys = this.getFieldKeys();
      //See if they are different (fields added for example)
      prevFieldKeys.sort();
      newFieldKeys.sort();
      if (JSON.stringify(prevFieldKeys) != JSON.stringify(newFieldKeys)) {
        //Query the lists again to get potentially new field data
        this.renderEvents();
      }
      return;
    }
    else if (prevProps.visJsonProperties != this.props.visJsonProperties) {
      reloadEvents = false;
      //Specify Timeline options
      try {
        let options = this.options;
        const userOptions = JSON.parse(this.props.visJsonProperties); //just in case
        options = this.extend(true, options, userOptions); //userOptions override set "defaults" above
        this._timeline.setOptions(options);
      }
      catch (e) {
        console.error(e);
      }
    }
    else if (prevProps.cssOverrides != this.props.cssOverrides) {
      this.renderDynamicStyles();
      return;
    }

    //Handle categories/legend & groups
    //this._timeline.setOptions({})
    this.renderDynamicStyles();
    this.renderLegend();
    //this.props.setGroups(this._timeline);
    this.setGroups();
    //Clear out bottom groups bar
    const bottomGroupsBar = document.getElementById("bottomGroupsBar-" + this.props.instanceId);
    if (bottomGroupsBar) {
      bottomGroupsBar.innerHTML = ""; //clear out then hide
      bottomGroupsBar.style.display = "none";
    }

    //Only re-render events if needed
    if (reloadEvents)
      this.renderEvents();
  }

  //Also fired after super.onPropertyPaneFieldChanged is called
  public render(): React.ReactElement<ITimelineCalendarProps> {
    //this.context == {} here
    //this.domElement == null
    //Had: ICustomDropdownOption, ICustomCollectionField
    const {instanceId} = this.props;

    return (
      <div>
        <div className={'container-' + this.props.instanceId}>
          <div id={"legend-" + instanceId} style={{display:"none"}} />
          <div id={"timeline-" + instanceId} />
        </div>
        <div id={"bottomGroupsBar-" + instanceId} className='bottomGroupsBar' />
        <div id={"dialog-" + instanceId} />
      </div>
    )
  }

  private filterTextForXSS(input:string): string {
    if (input == null)
      return "";

    //Filter to get Actual content from <html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><body><div>Actual content
    if (input.indexOf("<html>") !== -1) {
      const elem = document.createElement("div");
      elem.innerHTML = input; //removes html, head, and body tags but leaves <meta> tags
      //Filter out <meta> tags and trim to remove \n\n characters
      input = elem.innerHTML.replace(/<\/?meta*[^<>]*>/ig, '').trim();
    }
    
    //Escape additional elements besides just <script>
    //const whiteList = xss.getDefaultWhiteList(); //or xss.whiteList;
    input = input.replace(/javascript:/g, ''); //Extra protection for IE
    input = input.replace(/&#106;&#97;&#118;&#97;&#115;&#99;&#114;&#105;&#112;&#116;&#58;/g, '');
    input = input.replace(/&#0000106&#0000097&#0000118&#0000097&#0000115&#0000099&#0000114&#0000105&#0000112&#0000116&#0000058/g, '');
    input = input.replace(/&#x6A&#x61&#x76&#x61&#x73&#x63&#x72&#x69&#x70&#x74&#x3A/g, '');

    //Add SVG & other elements (no attributes, but that's handled later)
    whiteList.svg = ['xmlns', 'height', 'width', 'preserveaspectratio', 'viewbox', 'width', 'x', 'y'];
    whiteList.circle = ['cx', 'cy', 'r', 'pathlength'];
    whiteList.clippath = ['clippathunits'];
    whiteList.defs = [];
    whiteList.desc = [];
    whiteList.ellipse = ['cx', 'cy', 'rx', 'ry', 'pathlength'];
    whiteList.filter = ['x', 'y', 'width', 'height', 'filterunits', 'primitiveunits'];
    whiteList.foreignobject = ['x', 'y', 'width', 'height'];
    whiteList.g = [];
    whiteList.hatch = [];
    whiteList.hatchpath = [];
    whiteList.image = ['x', 'y', 'width', 'height', 'href', 'preserveaspectratio'];
    whiteList.line = ['x1', 'x2', 'y1', 'y2', 'pathlength'];
    whiteList.lineargradient = ['x1', 'x2', 'y1', 'y2', 'gradientunits', 'gradienttransform', 'href', 'spreadmethod'];
    whiteList.marker = ['markerheight', 'markerunits', 'markerwidth', 'orient', 'preserveaspectratio', 'refx', 'refy', 'viewbox'];
    whiteList.mask = ['height', 'maskcontentunits', 'maskunits', 'x', 'y', 'width'];
    whiteList.path = ['d', 'pathlength'];
    whiteList.pattern = ['height', 'href', 'patterncontentunits', 'patterntransform', 'patternunits', 'preserveaspectratio', 'viewbox', 'width', 'x', 'y'];
    whiteList.polygon = ['points', 'pathlength'];
    whiteList.polyline = ['points', 'pathlength'];
    whiteList.radialgradient = ['cx', 'cy', 'fr', 'fx', 'fy', 'gradientunits', 'gradienttransform', 'href', 'r', 'spreadmethod'];
    whiteList.rect = ['x', 'y', 'width', 'height', 'rx', 'ry', 'pathlength'];
    whiteList.set = ['to'];
    whiteList.stop = ['offset', 'stop-color', 'stop-opacity'];
    whiteList.switch = [];
    whiteList.symbol = ['height', 'preserveaspectratio', 'refx', 'refy','viewbox', 'width', 'x', 'y'];
    whiteList.text = ['x', 'y', 'dx', 'dy', 'rotate', 'lengthadjust', 'textlength'];
    whiteList.textpath = ['href', 'lengthaadjust', 'method', 'path', 'side', 'spacing', 'startoffset', 'textlength'];
    whiteList.title = [];
    whiteList.tspan = ['x', 'y', 'dx', 'dy', 'rotate', 'lengthadjust', 'textlength'];
    whiteList.use = ['href', 'x', 'y', 'width', 'height'];
    whiteList.view = ['viewbox', 'preserveaspectratio'];

    //Documentation: https://jsxss.com/en/options.html
    input = filterXSS(input, {
      whiteList: whiteList,
      stripIgnoreTagBody: true, //this would completely remove <iframe> vs. escaping it
      //attributes *not* in the whitelist for a tag
      onIgnoreTagAttr: function(tag:string, name:string, value:string, isWhiteAttr:boolean) {
        // If a string is returned, the value would be replaced with this string
        // If return nothing, then keep default (remove the attribute)
        //name is already lowercased
        //must return as full string: style="font-weight:bold"
        
        if (name === "id" || name === "style" || name === "class" || name === "title")
          return name + '="' + value + '"';
        
        if ((tag === "g" || tag === "hatch" || tag === "hatchpath") && name !== "onload")
          return name + '="' + value + '"';
      }
    });

    return input;
  }

  //Pass in the objects to merge as arguments (for a deep extend, set the first argument to true)
  //Cannot use Object.assign(options, userOptions) instead because it doesn't do deep property adding
  //TODO: Use this? https://developer.mozilla.org/en-US/docs/Web/API/structuredClone
  private extend(...args: any[]):any {
    const self = this;
    const extended = {} as any;
    let deep = false;
    let i = 0;
    const length = arguments.length;
  
    // Check if a deep merge
    if (Object.prototype.toString.call(arguments[0]) === '[object Boolean]') {
      deep = arguments[0];
      i++;
    }
  
    // Merge the object into the extended object
    const merge = function (obj:any):void {
      for (const prop in obj) {
        if (Object.prototype.hasOwnProperty.call(obj, prop)) {
          // If deep merge and property is an object, merge properties
          if (deep && Object.prototype.toString.call(obj[prop]) === '[object Object]') {
            extended[prop] = self.extend(true, extended[prop], obj[prop] );
          } else {
            extended[prop] = obj[prop];
          }
        }
      }
    };
  
    // Loop through each object and conduct a merge
    for ( ; i < length; i++ ) {
      const obj = arguments[i];
      merge(obj);
    }
  
    return extended;
  }

  private calcMaxHeight():number {
    const container = document.getElementById("legend-" + this.props.instanceId);
    const rect = container.getBoundingClientRect();
    //let win = element.ownerDocument.defaultView;
    //rect.top + win.pageYOffset (assuming page is scrolled down?)
    const height = window.innerHeight - rect.top - 295; //TODO: handle when "hide header & nav" clicked //was 140
    //this._timeline.setOptions({maxHeight: height});
    return height;

    //Real page is below, but doesn't exist in workbench
    //document.querySelector('div[data-automation-id="contentScrollRegion"]').offsetTop
    //document.querySelector('div[data-automation-id="contentScrollRegion"]').scrollTop == 0, not scrolled
  }

  //TODO: Rename since it's used for REST response data also - ISO format
  private formatDateFromSOAP(d:string):Date {//"2014-08-28 23:59:00" or if UTC "2019-12-13T05:00:00Z"
    if (d == null)
			return null;
		
    let theDate = null;

		//Check for UTC/Zulu time
		if (d.indexOf("Z") !== -1)
			theDate = new Date(d);
		else
			theDate = new Date(d.replace(" ", "T")); //needed for IE
    
    //Check for invalid date (from calculated columns for example)
    if (isNaN(theDate.getTime()))
      return null;
    else
      return theDate;
  }

  private options: any = { //TimelineOptions
    min: this.getMinDate(),      //lower limit of visible range
    max: this.getMaxDate(),         //upper limit of visible range
    zoomMin: 60000000, //allows zooming down to hour level (a smaller number further expands each hour)
    start: this.getViewStartDate(),       //initial start of loaded axis
    end: this.getViewEndDate(),           //initial end of loaded axis
    showCurrentTime: true,
    orientation: 'top',
    // orientation: {
    //   axis: "both",
    //   item: "top"
    // },
    minHeight: 50, //defaults to px
    //verticalScroll: true,
    //zoomKey: "ctrlKey",
    tooltip: {
      //followMouse: true,
      delay: 100,
      template: (item:any, elem:HTMLElement) => {
        if (this.props.tooltipEditor == null || this.props.tooltipEditor === "") {
          const handleTemplate = Handlebars.compile(this.props.getDefaultTooltip());
          const strResult = handleTemplate(item);
          return this.filterTextForXSS(strResult);
        }
        else {
          const handleTemplate = Handlebars.compile(this.props.tooltipEditor);
          const strResult = handleTemplate(item);
          return this.filterTextForXSS(strResult);
        }
      }
    },
    xss: {
      disabled: true //needed so that "class" can be used within tooltip (even when tooltip.template is used above)
    },
    groupEditable: true,
    groupOrder: function (a:any, b:any) {
      return a.order - b.order;
    },
    groupOrderSwap: function (a:any, b:any, groups:any) {
      const v = a.order;
      a.order = b.order;
      b.order = v;
    },
    order: function (a:any, b:any) {
      // Sort items vertically within groups by title (content field)
      // Items with no content will be sorted to the bottom
      const aContent = (a.content || "").toLowerCase();
      const bContent = (b.content || "").toLowerCase();
      if (aContent < bContent) return -1;
      if (aContent > bContent) return 1;
      return 0;
    }
    /*,
    groupTemplate: function (group, element) {
        if (!group) { return }
        ReactDOM.unmountComponentAtNode(element);
        return ReactDOM.render(<GroupTemplate group={group} />, element);
    }*/
  }

  private getViewStartDate(): Date {
    //Initial start view; default 7 days before today
    const viewStart = new Date();
    viewStart.setDate(viewStart.getDate() - (this.props.initialStartDays != null ? this.props.initialStartDays : 7)); //fix for 0 value
    return viewStart;
  }
  private getViewEndDate(): Date {
    //Set initial end view; default 3 months out
    const viewEnd = new Date();
    //viewEnd.setMonth(viewEnd.getMonth() + 3);
    viewEnd.setDate(viewEnd.getDate() + (this.props.initialEndDays || 90));
    return viewEnd;
  }

  private getMinDate(): Date {
    //Build dates for min/max data querying; default 2 months before today
    const minDate = new Date();
    minDate.setDate(minDate.getDate() - (this.props.minDays != null ? this.props.minDays : 60)); //fix for 0 value
    return minDate;
  }

  private getMaxDate(): Date {
    //Default max is 1 year from today; ensure the time is at the very end of the day (add a day but remove 1 second to get at the very end of the previous day)
    const now = new Date();
		const maxDate = new Date(now.getFullYear(), now.getMonth(), now.getDate()+1, 0, 0, -1);
    maxDate.setDate(maxDate.getDate() + (this.props.maxDays || 365));
    return maxDate;
  }
  
  private getGroupIdAtIndex(startingIndex:number): number {
    startingIndex = startingIndex || 0;
    let foundGroupId;

    const groups = this._dsGroups.get({
      order: "order",
      filter: function (item:any) {
        return (item.visible !== false);
      }
    });
    for (let i=0; i<groups.length; i++) {
      if (i === startingIndex) {
        foundGroupId = groups[i].id;
        break;
      }
    }

    return foundGroupId;
  }

  private initialBuildTimeline(): void {
    //Specify Timeline options
    //---------------------------------
    //First set any special
    if (this.props.calcMaxHeight)
      this.options.maxHeight = this.calcMaxHeight();

    //Incorporate any user provided overrides
    //let options = this.options;
    try {
      const userOptions = JSON.parse(this.props.visJsonProperties); //just in case
      this.options = this.extend(true, this.options, userOptions); //userOptions override set "defaults" above
    }
    catch (e) {
      console.error(e);
    }
    
    //Generate the Timeline
    //---------------------------------
    const container = document.getElementById("timeline-" + this.props.instanceId);
    this._dsGroups = new DataSet();
    this._dsItems = new DataSet();
    this._timeline = new Timeline(container, this._dsItems, this.options);
    
    //@ts-ignore
    window.TC = { //temp
       timeline: this._timeline,
       eventsDataSet: this._dsItems,
       groupsDataSet: this._dsGroups
    }

    //Add click handler for events
    this._timeline.on("select", (props) => {
      if (props.items.length > 0 && props.event.type === "tap") { //ignore the follow-on "press" event (still needed?)
				//Get the clicked on item object
        const oEvent = this._dsItems.get(props.items[0]);
        //Handle SP items
        if (oEvent.encodedAbsUrl) {
          const source = (oEvent.sourceObj as IListItem);
          const listConfigs = this.buildListConfigs(source);

          //Check for updates/deletion to item
          const checkSPOItem = () => {
            //TODO: Edits to series/recurring events with IDs like "16.0.2020-06-08T16:00:00Z" actually have a different ID generated by the SP form action
            /*
              SP item ID 6 created with recurrence pattern (fRecurrence & RecurrenceData fields)
              Instances of it's recurring events share the same EncodedAbsUrl '{site}/Lists/Calendar/6_.000'
              But each has it's own ID value: ID='6.0.2024-02-07T13:00:00Z'
                                              ID='6.0.2024-03-06T13:00:00Z'
            */
            if (oEvent.spId.toString().includes("T")) //or "-" or "Z"
				      return; //skip

            const fieldKeys = this.getFieldKeys();

            /* REST approach
            //TODO: Use the fieldKeys to get (and update) all relevant props
            //TODO: Add fAllDayEvent for calendar lists
            //TODO: Add encodedAbsUrl and others?
            const selectFields = "ID," + source.titleField + (source.isCalendar ? ",fAllDayEvent" : "");
              //(source.startDateField ? "," + source.startDateField + ",FieldValuesAsText/" + source.startDateField : "") + 
              //(source.endDateField ? "," + source.endDateField + ",FieldValuesAsText/" + source.endDateField : "");

            //Query for item
            this.props.context.spHttpClient.get(source.siteUrl + 
              //Previously had &$expand=FieldValuesAsText
              //Needed to expand user fields: select=*,Author/Title,Editor/Title&$expand=Author,Editor
              `/_api/web/lists/getById('${source.list}')/items?$filter=Id eq ${oEvent.spId}&$select=${selectFields}`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
              if (response.ok) {
                response.json().then((data:any) => {
                  if (data && data.value && data.value.length > 0) {
                    const item = data.value[0];
                    const updatedEvent = this.buildSPOItemObject(source, listConfigs, fieldKeys, item, oEvent.id);

                    //Add group (row/swimlane)
                    let multipleValuesFound = false;
                    if (listConfigs.groupId)
                      oEvent.group = listConfigs.groupId;
                    else if (listConfigs.groupField && this.props.groups) {
                      //Find the associated group to assign the item to
                      //ORIG: let groupFieldValue = elem.getAttribute("ows_" + listConfigs.groupField);
                      let groupFieldValue = item[listConfigs.groupField];
                      if (groupFieldValue) {
                        //ORIG: const groupSplit = this.handleMultiValues(groupFieldValue);
                        console.log(groupFieldValue); //TODO: check on object properties for lookups, users, etc.
                        const groupSplit: string[] = ["Test"];

                        //Look for "regular" values without ;# (they result in ["Single value"] array)
                        if (groupSplit.length == 1) {
                          groupFieldValue = groupSplit[0];
                        }
                        else {
                          //More than one, *real* value is in the field
                          multipleValuesFound = true;

                          //TODO: FIND any duplicate events and remove them

                          // //Create a duplicate event for each selected group value
                          // groupSplit.forEach(groupName => {
                          //   const eventClone = structuredClone(oEvent);
                          //   //Above duplicates the event object
                          //   eventClone.id = IdSvc.getNext(); //Set a new ID

                          //   //Find the associated group from it's name
                          //   this.props.groups.every((group:IGroupItem) => {
                          //     if (group.name == groupName) {
                          //       eventClone.group = group.uniqueId;
                          //       return false; //exit
                          //     }
                          //     else return true; //keep looping
                          //   });

                          //   //Add the clone to the DataSet
                          //   this._dsItems.add(eventClone);
                          // });
                        }

                        //Finalize single value events
                        if (multipleValuesFound == false) {
                          //Find the associated group from it's name
                          this.props.groups.every((group:IGroupItem) => {
                            if (group.name == groupFieldValue) {
                              oEvent.group = group.uniqueId;
                              return false; //exit
                            }
                            else return true; //keep looping
                          });
                        }
                      } //There is a groupFieldValue
                    } //A groupField was selected && there are this.props.groups

                    //Update the event
                    this._dsItems.update(updatedEvent);
                  }
                  else //Remove the item, it's been deleted
                    this._dsItems.remove(props.items[0]);
                });
              }
              else {
                //const statusCode = response.status;
                //const statusMessage = response.statusMessage; //May not exist?
                response.json().then(data => {
                  console.log(data);
                })
                .catch (error => {
                  //console.log("status: " + statusCode.toString() + " / " + statusNum.toString());
                  //reject("Error HTTP: " + response.status.toString() + " " + response.statusText);
                });
              }
            })
            .catch(error => {
              //console.log(error);
              //.message: "Unexpected end of JSON input"
              //.stack: "SyntaxError: Unexpected end of JSON input\n    at e.json..."
            });
            */

            //Use SOAP instead for easier handling of fields that might not exist (user inputted ones in the tooltip)
            let soapEnvelop = "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/'><soapenv:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" + 
              "<listName>" + source.list + "</listName>" + 
              "<viewFields><ViewFields>" + 
                //"<FieldRef Name='Title' />" + 
                (source.titleField ? "<FieldRef Name='" + source.titleField + "' />" : "") +
                //"<FieldRef Name='Location' />" + 
                "<FieldRef Name='EventDate' />" +
                "<FieldRef Name='EndDate' />" + 
                (source.isCalendar ? "<FieldRef Name='fRecurrence' />" : "") + 
                //(includeRecurrence ? "<FieldRef Name='RecurrenceData' />" : "") + 
                "<FieldRef Name='EncodedAbsUrl' />" +
                //fAllDayEvent seems to be included by default
                (listConfigs.classField ? "<FieldRef Name='" + listConfigs.classField + "' />" : "") +
                (listConfigs.groupField ? "<FieldRef Name='" + listConfigs.groupField + "' />" : "");

            //Add fields used in the tooltip
            fieldKeys.forEach(field => {
              soapEnvelop += "<FieldRef Name='" + field + "' />";
            });
            
            //Add date fields
            if (source.startDateField)
              soapEnvelop += "<FieldRef Name='" + source.startDateField + "' />";
            if (source.endDateField)
              soapEnvelop += "<FieldRef Name='" + source.endDateField + "' />";
            
            soapEnvelop += "</ViewFields></viewFields>" + 
              //rowLimit not needed since this is querying for just one item to update it's details
              "<query><Query><Where>" +
                "<Eq><FieldRef Name='ID'/><Value Type='Computed'>" + oEvent.spId.toString() + "</Value></Eq>" +
              "</Where></Query></query><queryOptions><QueryOptions>" + 
                //(includeRecurrence ? "<RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion>" : "") + 
                //(includeRecurrence ? "<ExpandRecurrence>TRUE</ExpandRecurrence>" : "") + 
                //(includeRecurrence ? "<RecurrenceOrderBy>TRUE</RecurrenceOrderBy>" : "") +
                "<ViewAttributes Scope='RecursiveAll' />" +
                //"<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" +
                "<ViewFieldsOnly>TRUE</ViewFieldsOnly>" +
                (listConfigs.dateInUtc === false ? "" : "<DateInUtc>TRUE</DateInUtc>") + //True returns dates as "2023-10-10T06:00:00Z" versus "2023-10-10 08:00:00"
              "</QueryOptions></queryOptions></GetListItems></soapenv:Body></soapenv:Envelope>";

            //Perform the query
            this.props.context.spHttpClient.post(source.siteUrl + "/_vti_bin/lists.asmx", SPHttpClient.configurations.v1,
            {
              headers: [
                ["Accept", "application/xml, text/xml, */*; q=0.01"],
                ["Content-Type", 'text/xml; charset="UTF-8"']
              ],
              body: soapEnvelop
            }).then((response: SPHttpClientResponse) => response.text())
            .then((strXml: any) => {
              // //Check for problems such as access denied to the site/web object
              // //They won't have an rs:data element with ItemCount attribute
              // if ($(jqxhr.responseXML).SPFilterNode("rs:data").attr("ItemCount") == null) {
              //   var strError = cal.listName + " returned invalid response";
              //     var isWarning = false;
              //     var isError = false;
              //     if (jqxhr.responseXML) {
              //         var msg = ($("title", jqxhr.responseXML).text() || "").trim();
              //         if (msg == "Access required") {
              //           isWarning = true;
              //           strError += ": Could not access site: " + cal.siteUrl;
              //         }
              //         else {
              //           strError += ": " + ($("h1.ms-core-pageTitle", jqxhr.responseXML).text() || "").trim();
              //           isError = true;
              //         }
              //     }
              //     else
              //       isError = true;
              //     TC.log(strError);
              //     dispatcher.queryCompleted(cal, isWarning, isError);
                  
              //   return; //don't proceed with the below
              // }
              
              //At this point we should have a valid list response
              const parser = new DOMParser();
              const xmlDoc = parser.parseFromString(strXml, "application/xml");
              xmlDoc.querySelectorAll("*").forEach(elem => {
                //Result when error happens: nodeName == "parsererror"

                if (elem.nodeName == "rs:data") {
                  if (elem.getAttribute("ItemCount") === "0")
                    //Remove the item, it's been deleted
                    this._dsItems.remove(props.items[0]);
                }
                else if (elem.nodeName === "z:row") { //actual data is here
                  //const itemDateInfo = this.getSPItemDates(source, listConfigs, elem);
                  const updatedEvent = this.buildSPOItemObject(source, listConfigs, fieldKeys, elem, oEvent.id);

                  //Add group (row/swimlane)
                  let multipleValuesFound = false;
                  if (listConfigs.groupId)
                    updatedEvent.group = listConfigs.groupId;
                  else if (listConfigs.groupField && this.props.groups) {
                    //Find the associated group to assign the item to
                    const groupFieldValue = elem.getAttribute("ows_" + listConfigs.groupField);
                    if (groupFieldValue) {
                      const groupSplit = this.handleMultipleSOAPValues(groupFieldValue);
                      
                      //OLD
                      // //Look for "regular" values without ;# (they result in ["Single value"] array)
                      // if (groupSplit.length == 1) {
                      //   groupFieldValue = groupSplit[0];
                      // }
                      // else {
                      
                      //Now need to assume that there *could* have been multiple selections previously
                      //but the item was updated to only have one (we need to remove those prior copies)

                      //More than one, *real* value is in the field
                      multipleValuesFound = true;

                      //Find duplicate events and remove them
                      const itemEvents = this._dsItems.get({
                        filter: function (item:any) {
                          return (item.spId === updatedEvent.spId); // && item.id != oEvent.id
                        }
                      });
                      this._dsItems.remove(itemEvents);

                      //Create a duplicate event for each selected group value
                      groupSplit.forEach(groupName => {
                        const eventClone = structuredClone(updatedEvent);
                        //Above duplicates the event object
                        eventClone.id = IdSvc.getNext(); //Set a new ID

                        //Find the associated group from it's name
                        this.props.groups.every((group:IGroupItem) => {
                          if (group.name === groupName) {
                            eventClone.group = group.uniqueId;
                            return false; //exit
                          }
                          else return true; //keep looping
                        });

                        //Add the clone to the DataSet
                        this._dsItems.add(eventClone);
                      });
                      //} //From OLD block above

                      /* Continuation from OLD block
                      //Finalize single value events
                      if (multipleValuesFound == false) {
                        //Find the associated group from it's name
                        this.props.groups.every((group:IGroupItem) => {
                          if (group.name == groupFieldValue) {
                            updatedEvent.group = group.uniqueId;
                            return false; //exit
                          }
                          else return true; //keep looping
                        });
                      }
                      */
                    } //There is a groupFieldValue
                  } //A groupField was selected && there are this.props.groups

                  //Update the event
                  if (multipleValuesFound === false)
                    this._dsItems.update(updatedEvent);
                }
              }); //xmlDoc.querySelectorAll

            }); //post.then
          };

          //Build URL to open clicked on item
					const itemUrl = oEvent.encodedAbsUrl.substring(0, oEvent.encodedAbsUrl.lastIndexOf("/")); //cut off the ending: "/ID#_.000"
          //objType 1 are  document library folders/doc sets and need /Forms in the URL
          let origUrl = itemUrl + (oEvent.objType === "1" ? "/Forms" : "") + "/DispForm.aspx?ID=" + oEvent.spId + "&Source=" + 
            encodeURIComponent(this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/inplview.aspx?Cmd=ClosePopUI");
          if (oEvent.objType === "1" && listConfigs.showFolderView)
            origUrl = oEvent.encodedAbsUrl;

          //Set initial "last"/previous URL
          let lastUrl = "about:blank";

          //Just open in new tab for now (until a custom modal dialog can be built that works for Modern & Classic pages)
          const winProxy = window.open(origUrl, "_blank");
          if (winProxy) { //Ensure browser didn't block opening the tab
            /*
            //For classic pages, need to overwrite the item deletion process
            winProxy.addEventListener("load", (e:Event) => {
              console.log("load: " + winProxy.location.href);
              const winObj = winProxy as any;
              if (winObj._ChangeLayoutMode || winObj._EditItem) {
                const orig_DeleteItemConfirmation = winObj.DeleteTimeConfirmation;
                winObj.DeleteItemConfirmation = function() {
                  var result = orig_DeleteItemConfirmation();
                  if (result) {
                    const form = winProxy.document.getElementById("aspnetForm");
                    const baseUrl = form.getAttribute("action");
                    //Overwrite the action with a Source to force the dialog closed
                    form.setAttribute("action", baseUrl + "&Source=" + encodeURIComponent(winObj._spPageContextInfo.webAbsoluteUrl + "/_layouts/15/inplview.aspx?Cmd=ClosePopUI"));
                  }
                  return result;
                }
              }
            }, false);
            */
            winProxy.focus(); //just to make sure
            
            //Interval timer needed to *consistently* and *repeatedly* check if page has changed or win closed
            const timerId = setInterval((handler:TimerHandler) => {
              //winProxy.document and .location never seem to be null, so check .closed prop
              if (winProxy.closed != true) {
                //Has the special "close page" been reached?
                //@ts-ignore @typescript-eslint/TS2550 (for endsWith)
                if (winProxy.location.pathname.endsWith("/inplview.aspx") && winProxy.location.search === "?Cmd=ClosePopUI") {
                  //console.log("found Cmd=ClosePopUI, closing window");
                  clearInterval(timerId);
                  window.focus(); //doesn't always focus back on main window (if back button is clicked)
                  winProxy.close();
                  checkSPOItem();
                  return;
                }

                //Check if page was navigated
                if (winProxy.location.href !== lastUrl) {
                  //Check for classic vs modern page (could pre-check if not a known classic calendar)
                  //@ts-ignore
                  if (winProxy._ChangeLayoutMode || winProxy._EditItem) {
                    //This should already follow the Source redirect option to the inplview.aspx page
                  }
                  else {
                    //Modern page (doesn't properly keep Source param from Disp to EditForm), so check if we are back at the original/DispForm page
                    if (winProxy.location.href === origUrl && lastUrl !== "about:blank") {
                      //console.log("force close the window here!");
                      clearInterval(timerId);
                      window.focus(); //doesn't always focus back on main window (if back button is clicked)
                      winProxy.close();
                      checkSPOItem();
                      return;
                    }
                  }

                  //console.log("Interval: URL was changed: " + lastUrl);
                  lastUrl = winProxy.location.href;
 
                  //Pointless to add a "load" handler here as it fires right away since this is a new page/URL change
                  
                  // winProxy.addEventListener("beforeunload", (e:Event) => {
                  //   //console.log("page changed-> beforeunload: " + winProxy.location.href);
                  //   winProxy.opener.someFunction("inside interval-> beforeunload: " + winProxy.location.href);
                  // }, false);
                }
              }
              else {
                //console.log("winProxy null or winProxy.closed");
                clearInterval(timerId);
                checkSPOItem(); //needed for case where single field updated in Modern page and user closed tab
                window.focus(); //may not help at all
              }
            }, 250);
          }

          //NOTE: "readystatechange" and "navigate" and "popstate" never fired
          /*fired before unload when page is closed (but you cannot tell the difference):
          //  visibilitychange (hidden) (winProxy.closed: false) -> but it's same as first firing for about:blank
          winProxy.addEventListener("visibilitychange", (e:Event) => {
            console.log("visibilitychange (" + winProxy.document.visibilityState + ") (winProxy.closed: " + winProxy.closed + "): " + winProxy.location.href);
            
            //adding load here after about:blank doesn't help
          }, false);
          */
          //visibilitychange
          //  console.log("visibilitychange (" + winProxy.document.visibilityState + "): " + winProxy.location.href);
          //  visibilitychange (hidden) also for when page is about to unload (while it is *visible*)
          //DOMContentLoaded once
          //Load fired here once
          //Intervals check here
          //beforeunload (when page is navigated; not fired when tab is closed)
          //  & window.someFunction called!
          //visibilitychange (same DispForm.aspx page)
          
          //unload (same DispForm.aspx page)
          /* deprecated (but fired when page is closed) -> first for about:blank
          winProxy.addEventListener("unload", (e:Event) => {
            console.log("unload (deprecated): " + (winProxy && winProxy.location && winProxy.location.href || "null"));

            //fired only for first page after about:blank unloaded
            winProxy.addEventListener("load", (e:Event) => {
              console.log("load from unload: " + winProxy.location.href);
            }, false);
          }, false);
          */

          return; //Below could be used but #s4-workspace class .ms-core-overlay (for classic) needs to have
                //style manually set: height: 714px; width: 592px; overflow-y: auto;

          const dlgContainer = document.getElementById("dialog-" + this.props.instanceId);
          const element1: React.ReactElement<IFrameDialogProps> = React.createElement(IFrameDialog,
            {
              url: itemUrl + "/DispForm.aspx?ID=" + oEvent.spId, //+ IsDlg=1 for classic
              iframeOnLoad: (iframe: any) => {
                  console.log("iframe loaded");
                },
              hidden: false,
              onDismiss: (event: React.MouseEvent) => {
                console.log("dialog dismissed");
                ReactDom.unmountComponentAtNode(dlgContainer);
              },
              modalProps: {
                isBlocking: true,
                //containerClassName: styles.dialogContainer
              },
              dialogContentProps: {
                type: DialogType.close,
                showCloseButton: true
              },
              width: "600px",
              height: "315px"
            }
          );
          ReactDom.render(element1, dlgContainer);
        }

        //Handle Outlook events
        if (oEvent.calEventWebLink) {
          //Just open in new tab for now
          //window.open(oEvent.calEventWebLink, "_blank");

          //Just open in new tab for now (until a custom modal dialog can be built that works with cross-origin)
          const calProxy = window.open(oEvent.calEventWebLink, "_blank");
          if (calProxy) { //Ensure browser didn't block opening the tab
            calProxy.focus(); //just to make sure

            //Cannot use an interval here to check calProxy.closed as it *always* shows true, despite being allowed per:
            //https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy#window

            //Look for when the timeline tab/window is focused on again
            // const handleVisibilityChange = function() {
            //   console.log("handleVisibilityChange, hidden:" + document.hidden + ", visibilityState:" + document.visibilityState);
            // }
            // document.addEventListener("visibilitychange", handleVisibilityChange);
            const timelineFocused = ():void => {
              //Remove listener
              window.removeEventListener('focus', timelineFocused);

              //Query for item and update in Timeline
              const source = (oEvent.sourceObj as ICalendarItem);
              const calConfigs = this.buildCalendarConfigs(source);
              this.queryCalendar(source, calConfigs, oEvent, 0);
            }
            window.addEventListener('focus', timelineFocused);
          }
        }
      }
    });

    this._timeline.on('doubleClick', (props) => {
      //TODO: Complete for adding events to calendar(s)
      //properties.group: 4
      //properties.time: date/time wehre clicked
      //properties.snappedTime._d for start of day clicked and ._i for same as .time

      //addEvent(props)
      //  opens pop up page and then queries for item detail to dynamically add to timeline
    });

    //Build references
    const labelSetElem = document.querySelector("#timeline-" + this.props.instanceId + " .vis-labelset");
    const bottomGroupsBar = document.getElementById("bottomGroupsBar-" + this.props.instanceId);

    function GetVisLabelElement(elem:HTMLElement): HTMLElement {
      if (elem.classList.contains("vis-labelset")) //we've gone too far up the dom
        return null; //shouldn't happen
      else if (elem.classList.contains('vis-label'))
        return elem; //found the element
      else
        return GetVisLabelElement(elem.parentElement); //try it's parent
    }

    //Add click handler for group click action to "remove" group
    //TODO: Support use of "dblclick" instead?
    //Could further refine starting with div.vis-left
    labelSetElem.addEventListener('click', (e:Event) => {
      const labelElem = GetVisLabelElement(e.target as HTMLElement);
      //@ts-ignore //Array.from is valid
      const elemIndex = Array.from(labelElem.parentElement.children).indexOf(labelElem);

      //Find the selected group
      const foundGroupId = this.getGroupIdAtIndex(elemIndex);
      //Hide the selected group
      const theGroup = this._dsGroups.get(foundGroupId);
      this._dsGroups.updateOnly({id: theGroup.id, visible:false});
      
      //Show the bar
      bottomGroupsBar.style.display = "block";
      
      //Add to bottom bar
      const newElem = document.createElement("div");
      newElem.className = "vis-item vis-range";
      newElem.dataset.groupObj = JSON.stringify(theGroup);
      newElem.innerText = theGroup.name;
      bottomGroupsBar.append(newElem);
    });

    //Add right-click handler to hide all other groups
    labelSetElem.addEventListener('contextmenu', (e:Event) => {
      const labelElem = GetVisLabelElement(e.target as HTMLElement);
      if (labelElem) {
        e.preventDefault(); //stop the normal menu from appearing

        //Is this the only group currently being shown?
        const shownGroups = this._dsGroups.get({
          filter: function (group:any) {
            return (group.visible !== false);
          }
        });

        if (shownGroups.length === 1) {
          //Process each group boxes
          bottomGroupsBar.querySelectorAll("div.vis-item").forEach((groupElem:HTMLElement) => {
            processBottomBarItem(groupElem);
          })
          bottomGroupsBar.style.display = "none";
          return; //exit
        }

        //Find the selected group
        //@ts-ignore //Array.from is valid
        const elemIndex = Array.from(labelElem.parentElement.children).indexOf(labelElem);
        const foundGroupId = this.getGroupIdAtIndex(elemIndex);

        //Get groups to remove
        const groupsToRemove = this._dsGroups.get({
          filter: function (group:any) {
            return (group.id !== foundGroupId && group.visible !== false);
          }
        });
        
        //Show bar
        bottomGroupsBar.style.display = "block";
        
        //Add groups to bar
        groupsToRemove.forEach((group:any) => {
          //Add to bottom bar
          const newElem = document.createElement("div");
          newElem.className = "vis-item vis-range";
          newElem.dataset.groupObj = JSON.stringify(group);
          newElem.innerText = group.name;
          bottomGroupsBar.append(newElem);
          
          //Set visible property
          this._dsGroups.updateOnly({id: group.id, visible:false});
        });
      }
    });

    //Add click handler for bottomGroupsBar boxes
    bottomGroupsBar.addEventListener('click', (e:Event) => {
      //Only process when an actual group box is clicked
      const childElem = e.target as HTMLDivElement;
      if (childElem.classList.contains('vis-item')) {
        processBottomBarItem(childElem);
        
        //Are there any group boxes still inside the bottomBar?
        if (bottomGroupsBar.querySelectorAll(".vis-item").length === 0)
          bottomGroupsBar.style.display = "none";
      }
    });

    //Function: processBottomBarItem
    const processBottomBarItem = (elem:HTMLElement):any => {
      const groupToAdd = JSON.parse(elem.dataset.groupObj);
      //Remove the group box from inside the bottomGroupsBar
      elem.remove();
      
      //Make visible again within timeline
      this._dsGroups.updateOnly({id: groupToAdd.id, visible:true});
      
      return groupToAdd;
    }

    //Add weekend items
    if (true) {//Future: this.props.shadeWeekends) {
      let weekendStart = null as Date;
      if (this.options.min.getDay() === 0) { //0 = Sunday
        //Need to set weekend one day before
        weekendStart = new Date(this.options.min.getFullYear(), this.options.min.getMonth(), this.options.min.getDate() - 1);
      }
      else if (this.options.min.getDay() === 6) { //6 = Saturday
        weekendStart = new Date(this.options.min.getFullYear(), this.options.min.getMonth(), this.options.min.getDate());
      }
      else {
        //Add days to get to Saturday
        const daysToAdd = 6 - this.options.min.getDay();
        weekendStart = new Date(this.options.min.getFullYear(), this.options.min.getMonth(), this.options.min.getDate() + daysToAdd);
      }
      //Add two days but remove 1 second to get at the very end of the previous day
      const weekendEnd = new Date(weekendStart.getFullYear(), weekendStart.getMonth(), weekendStart.getDate()+2, 0, 0, -1);
      
      //Generate all weekends
      while (weekendStart < this.options.max) {
        const oEvent = {
          id: IdSvc.getNext(),
          content: "",
          start: weekendStart,
          end: weekendEnd,
          type: "background",
          className: "weekend"
        };
        //Add to dataset for immedate rendering
        this._dsItems.add(oEvent);
        
        //Get next weekend
        weekendStart.setDate(weekendStart.getDate() + 7);
        weekendEnd.setDate(weekendEnd.getDate() + 7);
      }
    }

    this.renderDynamicStyles();
    this.renderLegend(true);
    this.setGroups();
    this.renderEvents();
  }

  private renderDynamicStyles(): void {
    const styleId = "TimelineDynStyles-" + this.props.instanceId.substring(24); //use last portion of GUID
    let styleElem = document.getElementById(styleId);
    if (styleElem === null) {
      //Add the styles
      const head = this.props.domElement.ownerDocument.head;
      styleElem = document.createElement("style");
      //styleElem.type = 'text/css';
      styleElem.id = styleId;
      //SPFx default styles have attr: data-load-themed-styles="true"

      //was setting inner here
      head.appendChild(styleElem);
    }

    //Set default
    let styleHtml = `.container-${this.props.instanceId} {
  width: 100%;
  height: 100%;
}
`;
    
    //Update the Category css <style>
    //let styleHtml = '';
    if (this.props.categories) {
      this.props.categories.forEach((categoryItem: ICategoryItem) => {
        //Example: '.cssClass { color: #f00; }'
        //Adding .vis-item before .cssClass prevents .vis-item.vis-selected class from applying yellow selection border
        // styleHtml += '.vis-item.' + this.ensureValidClassName(value.name) + ' { ' + //was starting with padding:0 5px; border-radius:2px; 
        //   'border-color:' + value.borderColor + '; color:' + value.textColor + ' ' +
        //   (value.bgColor && value.bgColor != '' ? 'background-color:' + value.bgColor : '' ) + '}\r\n'; //border:1px solid ' + value.borderColor

        let divStyles = this.props.buildDivStyles(categoryItem);
        if (this.props.hideItemBoxBorder && divStyles.indexOf("background-color") === -1) //set bg to border color, mostly only applicable for vertical Holidays
          divStyles += "background-color:" + categoryItem.borderColor + ";";
        styleHtml += '.vis-item.' + this.props.ensureValidClassName(categoryItem.name) + ' {' + divStyles + '}\r\n';

        //Old way below when "advancedStyles" were JSON based
        //-----------------------------------
        // //Convert JavaScript styles object into "cssText" format
        // //const divStyles = this.props.buildDivStyles(categoryItem);
        // if (divStyles.backgroundColor == null)
        //   divStyles.backgroundColor = categoryItem.borderColor; //set background to border, mostly only applicable for vertical Holidays
        // const newElem = document.createElement("div")
        // for (const key in divStyles) {
        //   // @ts-ignore
        //   newElem.style[key] = divStyles[key];
        // }
        // styleHtml += '.vis-item.' + this.ensureValidClassName(categoryItem.name) + ' {' + newElem.style.cssText + '}\r\n';
      });
    }

    if (this.props.hideItemBoxBorder)
    styleHtml += `
  /* Use a thin line instead of a colored box for all multi-day events */
  /* .vis-timeline ----------------------------------------------------------------- */
  .vis-item.vis-range, .vis-item.legendBox {
      background-color: transparent;
      border-style: none none solid none;
      border-bottom-width: 7px;
      border-radius: 0px;
  }
  .vis-item.vis-range .vis-item-content {
      padding:0px;
  }
  /* ----------------------------------------------------------------- */

`;

  if (this.props.overflowTextVisible)
  styleHtml += `
  /* If you want the full text of multi-day events to display even if it goes outside of their box, use this CSS below */
  .vis-item.vis-range .vis-item-overflow {
      overflow: visible;
  }

`;

    if (this.props.fillFullWidth)
      styleHtml += `
  /* Force full width of page, and for workbench environment too */
  #SPPageChrome section.mainContent div.SPCanvas div.CanvasZone[data-automation-id="CanvasZone"] > div:first-child,
  #SPPageChrome section.mainContent div.SPCanvas div.CanvasZone[data-automation-id="CanvasZone"] > div:first-child div[data-automation-id="CollapsibleLayer-Content"],
  #workbenchPageContent div.CanvasComponent div.Canvas > div.CanvasZoneContainer > div.CanvasZone:first-child {
    max-width:unset;
  }
  `;

  //   if (this.props.hideSocialBar)
  //     styleHtml += `
  // /* Hide social/comments box at page bottom */
  // #SPPageChrome section.mainContent #CommentsWrapper {
  //   display:none;
  // }
  // `;

    if (this.props.cssOverrides) {
      let cssOverrides = this.props.cssOverrides.replace(/javascript:/g, ''); //Extra protection for IE
      cssOverrides = filterXSS(cssOverrides);

      styleHtml += "/* CSS class overrides */\r\n" + cssOverrides;
    }

    //Add override for default holiday class
    /* Special event items */
    /*.vis-item.Holiday, .vis-item.holiday {
      background-color: rgba(255, 255, 0, .2);
      border-color: rgba(255, 255, 0, .3);
    }*/

    styleElem.innerHTML = styleHtml;
  }

  private renderLegend(initial?:boolean): void {
    //Add the visual legend boxes
    const legend = document.getElementById("legend-" + this.props.instanceId);
    if (legend && this.props.categories) {
      legend.innerHTML = ''; //TODO: https://stackoverflow.com/questions/3955229/remove-all-child-elements-of-a-dom-node-in-javascript
      this.props.categories.forEach((value: ICategoryItem) => {
        if (value.visible) {
          const newElem = document.createElement("div");
          newElem.className = "legendBox vis-item vis-range " + this.props.ensureValidClassName(value.name);
          newElem.dataset.className = this.props.ensureValidClassName(value.name);
          newElem.innerText = value.name;
          legend.appendChild(newElem);
        }
      });
    }

    if (initial) {
      //Add the loading image
      const newElem = document.createElement("div");
      newElem.id = "loading-" + this.props.instanceId;
      newElem.innerHTML = "<img src='/_layouts/images/kpiprogressbar.gif' />";
      legend.before(newElem);
    }

    //TODO: Move the below up into the "initial" block so it's not executed multiple times?

    //Add client event to legend boxes
    const legendBar = document.getElementById("legend-" + this.props.instanceId);
    document.querySelectorAll('#legend-' + this.props.instanceId + ' > .legendBox').forEach(el => {
      el.addEventListener('click', event => {
        //Ignore print button
        // if ($(this).hasClass("print"))
        //   return;
        
        //Get the groupId
        //var groupId = $(this).attr("data-groupId"); //will not be defined for categories

        //Is the category to be shown?
        if (el.classList && el.classList.contains('gray')) {
          //Add the data back in
          this._dsItems.add(JSON.parse((el as HTMLElement).dataset.events));
        }
        else { //Hide data
          const className = (el as HTMLElement).dataset.className;//getAttribute("data-className");
          const itemEvents = this._dsItems.get({
            filter: function (item:any) {
              //return (item.className == className);
              //Handle multiple classes such as "Meeting Pending"
              return (item.className && item.className.split && item.className.split(" ").indexOf(className) !== -1);
            }
          });
          (el as HTMLElement).dataset.events = JSON.stringify(itemEvents);
          this._dsItems.remove(itemEvents);
        }

        el.classList.toggle('gray');
      });
    });

    //Add right-click handler to legend boxes
    legendBar.addEventListener('contextmenu', (e:Event) => {
      //Only process when an item was actually clicked
      const childElem = e.target as HTMLDivElement;
      if (childElem.classList.contains('legendBox')) {
        e.preventDefault(); //stop the normal menu from appearing
        
        //Is this the only one being shown?
				//var $legendBoxes = $("#legend > .legendBox:not(.gray):not(.print)");
        const activeLegendBoxes = legendBar.querySelectorAll(".legendBox:not(.gray)");
        if (activeLegendBoxes.length === 1 && childElem.classList.contains("gray") === false) {
          //Restore all inactive categories
          legendBar.querySelectorAll(".legendBox.gray").forEach((elem:HTMLElement) => {
            elem.click();
          });

          return; //exit
        }
        
        //Is this box currently grayed out?
        if (childElem.classList.contains("gray")) {
          //Activate the category
          childElem.click();
        }
        
        //Hide all other ACTIVE categories
        activeLegendBoxes.forEach((elem:HTMLElement) => {
          if (childElem !== elem)
            elem.click();
        });
      }
    });
  }

  private setGroups(): void {
    if (this.props.groups) {
      //Reformat array
      const groupsFormatted = this.props.groups.map((g:IGroupItem) => {
        let theContent = "";
        if (g.html) {
          const handleTemplate = Handlebars.compile(g.html);
          theContent = handleTemplate(g);
          theContent = this.filterTextForXSS(theContent);
        }
        return {
          id: g.uniqueId,
          content: (g.html ? theContent : g.name),
          name: g.name, //For use in the bottomGroupsBar
          order: g.sortIdx,
          visible: g.visible,
          className: g.className
        }
      });

      this._dsGroups.clear();
      if (groupsFormatted.length === 0)
        this._timeline.setGroups(null); //needed to get events to render if there are no groups
      else {
        this._dsGroups.add(groupsFormatted);
        this._timeline.setGroups(this._dsGroups);
      }
    }
  }

  private renderEvents(): void {
    //Prevent concurrent event loading to avoid duplicates
    if (this._isLoadingEvents) {
      console.log("TimelineCalendar: Skipping renderEvents() - already loading");
      return;
    }
    console.log("TimelineCalendar: Starting renderEvents()");
    this._isLoadingEvents = true;

    //Function: showLegend (must be delared before/above where it's called)
    const showLegend = ():void => {
      //Hide the loader image
      document.getElementById("loading-" + this.props.instanceId).style.display = "none";
      //Show the legend
      document.getElementById("legend-" + this.props.instanceId).style.display = "block";
    }

    //Remove any existing events to prevent duplicate event adding (while in edit mode)
    const itemEvents = this._dsItems.get({
      filter: function (item:any):boolean {
        return (item.className !== "weekend");
      }
    });
    this._dsItems.remove(itemEvents);

    //Get SharePoint list/calendar data
    let spPromise = null as Promise<void | any[]>;
    if (this.props.lists) {
      //Get the view CAML
      spPromise = this.getViewsCAML().then(() =>{
        //Now get the events
        return this.getSharePointEvents().then(response => {
          //console.log("all data returned, response is undefined because no data is actually returned");
        });
      });
    }

    //Get Outlook calendar events
    let outlookPromise = null as Promise<void | any[]>;
    if (this.props.calsAndPlans) {
      outlookPromise = this.getOutlookEvents();
    }

    //When both are finished
    Promise.all([spPromise, outlookPromise]).then(response => {
      console.log("TimelineCalendar: Events loaded successfully");
      showLegend();
      this._isLoadingEvents = false;
    }).catch(error => {
      //Ensure flag is cleared even on error
      console.error("TimelineCalendar: Error loading events:", error);
      this._isLoadingEvents = false;
      showLegend();
    });
  }

  private getViewsCAML(): Promise<any[]> {
    //Build a new array of unique calendars to avoid querying the same one multiple times
    return Promise.all(this.props.lists.map((list:IListItem) => {
      //Only for lists that have View specified
      if (list.view !== null && list.view.trim() !== "") {
        //Build list filter, first assume a title then check if GUID
        let listFilter = "lists/getByTitle('" + list.list + "')";
        const guidRegex = /[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}/i;
        if (guidRegex.test(list.list))
          listFilter = "lists(guid'" + guidRegex.exec(list.list)[0] + "')";
        
        //Get ViewQuery
        return this.props.context.spHttpClient.get(list.siteUrl + "/_api/web/" + listFilter + "/views?$select=Id,Title,ViewQuery", 
          SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => response.json())
          .then((data:any) => {
            if (data == null || data.value == null) //user may have no access to this list
              return;

            data.value.forEach((view:any) => {
              if (list.view.toLowerCase() === view.Title.toLowerCase() || list.view.toLowerCase().indexOf(view.Id) !== -1) {
								//TC.log("Got ViewQuery for '" + view.Title + "' view");
								//Legacy calendar view (view.ViewType: "CALENDAR") & Standard w/ Recurrence:
                //"<Where><And><DateRangesOverlap><FieldRef Name="EventDate" /><FieldRef Name="EndDate" /><FieldRef Name="RecurrenceID" /><Value Type="DateTime"><Month /></Value></DateRangesOverlap><Eq><FieldRef Name="Category" /><Value Type="Text">Birthday</Value></Eq></And></Where>"
                //Modern list calendar view (view.ViewType: "HTML" && view.ViewType2: "MODERNCALENDAR"):
                //"<Where><DateRangesOverlap><FieldRef Name=\"StartDate\" /><FieldRef Name=\"EndDate\" /><Value Type=\"DateTime\"><Month /></Value></DateRangesOverlap></Where>"
								//Could also be blank "" or "<Where><Eq><FieldRef Name=\"MultiChoice\" /><Value Type=\"Text\">Category 1</Value></Eq></Where>"
                //And <OrderBy><FieldRef Name=\"ID\" /></OrderBy> could be before <Where> as part of the ViewQuery

								//Find DateRangeOverlap & Where positions
								const droStartIndex = view.ViewQuery.indexOf('<DateRangesOverlap>');
                const droEndIndex = view.ViewQuery.indexOf('</DateRangesOverlap>');
								const andWhereEndIndex = view.ViewQuery.lastIndexOf('</And></Where>');
								const whereStartIndex = view.ViewQuery.indexOf('<Where>');
								const whereEndIndex = view.ViewQuery.lastIndexOf('</Where>');
                
                //Ignore DateRangeOverlap & extract just the actual filter portion
								if (droEndIndex > -1 && andWhereEndIndex > -1) //legacy calendar views
									list.viewFilter = view.ViewQuery.substring(droEndIndex+20, andWhereEndIndex);
                else if (droStartIndex > -1 && droEndIndex > -1) { //modern calendar views
                  const noDroWhereQuery = view.ViewQuery.substring(0, droStartIndex) + view.ViewQuery.substring(droEndIndex+20)
                  if (noDroWhereQuery == "<Where></Where>")
                    list.viewFilter = "";
                  else //Just get the query inside the <Where>
                    list.viewFilter = view.ViewQuery.substring(whereStartIndex+7, whereEndIndex);
                }
                //Just get the query inside the <Where>
								else if (whereStartIndex > -1)
									list.viewFilter = view.ViewQuery.substring(whereStartIndex+7, whereEndIndex);
                //No <Where> so ensure a blank filter is set in case user switched from another view
                else
                  list.viewFilter = "";
								
                //Examples
                //-------------------------
								//Real Calendar
								//View Query: "<Where><And><DateRangesOverlap><FieldRef Name=\"EventDate\" /><FieldRef Name=\"EndDate\" /><FieldRef Name=\"RecurrenceID\" /><Value Type=\"DateTime\"><Year /></Value></DateRangesOverlap><Eq><FieldRef Name=\"Category\" /><Value Type=\"Text\">Birthday</Value></Eq></And></Where>"
								//ViewData: "<FieldRef Name=\"Title\" Type=\"CalendarMonthTitle\" /><FieldRef Name=\"Title\" Type=\"CalendarWeekTitle\" /><FieldRef Name=\"Location\" Type=\"CalendarWeekLocation\" /><FieldRef Name=\"Title\" Type=\"CalendarDayTitle\" /><FieldRef Name=\"Location\" Type=\"CalendarDayLocation\" />"
								
								//Task list with Calendar view
								//ViewQuery: "<Where><DateRangesOverlap><FieldRef Name=\"StartDate\" /><FieldRef Name=\"DueDate\" /><Value Type=\"DateTime\"><Month /></Value></DateRangesOverlap></Where>"
								//ViewData: "<FieldRef Name=\"Title\" Type=\"CalendarMonthTitle\" /><FieldRef Name=\"Title\" Type=\"CalendarWeekTitle\" /><FieldRef Name=\"Title\" Type=\"CalendarWeekLocation\" /><FieldRef Name=\"Title\" Type=\"CalendarDayTitle\" /><FieldRef Name=\"\" Type=\"CalendarDayLocation\" />"
							}
            })
          })
          .catch((error:any) => {
            console.error(error);
          });
      }
    }))
  }

  private getFieldKeys(prevValue?:string):string[] {
    let fieldKeys = [] as string[];
    const source = (prevValue || this.props.tooltipEditor);
    if (source) {
      //Extract the {{property}} references which includes {{{triple references}}}
      fieldKeys = source.match(/{{(.*?)}}/g);
      if (fieldKeys) {
        //Extract just the field/"property" text from inside the {{ }} or {{{ }}}
        for (let i=0; i < fieldKeys.length; i++) {
            const matchResults = fieldKeys[i].match(/\w+/g);
            if (matchResults.length === 1)
              fieldKeys[i] = matchResults[0];
            else //handle cases like "{{{limit Description}}}" where the actual Description property is at the end of the match
              fieldKeys[i] = matchResults[1];
        }

        //Remove the vis.js "default" fields
        fieldKeys = fieldKeys.filter(i => {
          if (i !== "content" && i !== "start" && i !== "end")
              return i;
        });
      }
    }
    return fieldKeys;
  }

  private buildListConfigs(list:IListItem): IListConfigs {
    let configs = {} as IListConfigs;
    if (list.configs && list.configs.trim() !== "") {
      try {
        configs = JSON.parse(list.configs);
      }
      catch (e) {
        //Nothing needed
      }
    }

    //Ensure properties have valid types (and set default values if needed)
    if (configs.camlFilter == null || typeof configs.camlFilter !== "string")
      configs.camlFilter = null;
    if (configs.dateInUtc == null || typeof configs.dateInUtc !== "boolean")
      configs.dateInUtc = true;
    if (configs.visible == null || typeof configs.visible !== "boolean")
      configs.visible = true;
    if (configs.multipleCategories == null || typeof configs.multipleCategories !== "string")
      configs.multipleCategories = "useFirst";
    if (configs.extendEndTimeAllDay == null || typeof configs.extendEndTimeAllDay !== "boolean")
      configs.extendEndTimeAllDay = true;
    if (configs.fieldValueMappings == null || typeof configs.fieldValueMappings !== "object") //null is an "object"
        configs.fieldValueMappings = {};
    if (configs.showFolderView == null || typeof configs.showFolderView !== "boolean")
        configs.showFolderView = false;
    if (configs.limitHolidayToRow == null || typeof configs.limitHolidayToRow !== "boolean")
        configs.limitHolidayToRow = false;
    //When adding new props, consider the effects of the prop *not* being provided/set at all

    //Add Category props to configs (classField and className)
    //Split on the : char to determine if a field or category was selected (Field:fieldInternalName or Static:category.uniqueId)
    if (list.category) {
      const catValues = list.category.split(":");
      if (catValues[0] === "Field") {
        configs.classField = catValues[1];
        if (configs.className)
          configs.className = null;
      }
      else { //catValues[0] assumed to be "Static"
        const categoryId = catValues[1]; //Will be the uniqueId, need to get the display name next
        if (this.props.categories) {
          this.props.categories.every((category:ICategoryItem) => {
            if (category.uniqueId === categoryId) {
              configs.className = category.name; //store the display name instead
              if (configs.classField)
                configs.classField = null;
              return false; //exit
            }
            else return true; //keep looping
          });
        }
      }
    }

    //Add Group/Row props to configs (groupField and groupId)
    //Split on the : char to determine if a field or category was selected
    if (list.group) {
      const catValues = list.group.split(":");
      if (catValues[0] === "Field") {
        configs.groupField = catValues[1];
        if (configs.groupId)
          configs.groupId = null;
      }
      else { //[0] assumed to be "Static"
        configs.groupId = catValues[1]; //Will be the uniqueId
        if (configs.groupField)
          configs.groupField = null;
      }
    }

    return configs;
  }

  private buildCalendarConfigs(calendar:ICalendarItem): ICalendarConfigs {
    let calConfigs = {} as ICalendarConfigs;
      if (calendar.configs && calendar.configs.trim() !== "") {
        try {
          calConfigs = JSON.parse(calendar.configs);
        }
        catch (e) {
          //Nothing needed
        }
      }

      //Ensure properties have valid types (and set default values if needed)
      if (calConfigs.visible == null || typeof calConfigs.visible !== "boolean")
        calConfigs.visible = true;
      if (calConfigs.multipleCategories == null || typeof calConfigs.multipleCategories !== "string")
        calConfigs.multipleCategories = "useFirst";
      if (calConfigs.fieldValueMappings == null || typeof calConfigs.fieldValueMappings !== "object") //null is an "object"
        calConfigs.fieldValueMappings = {};
      //When adding new props, consider the effects of the prop *not* being provided/set at all

      //Add Category props to configs (classField and className)
      //Split on the : char to determine if a field or category was selected (Field:owaField or Static:category.uniqueId)
      if (calendar.category) {
        const catValues = calendar.category.split(":");
        if (catValues[0] == "Field") {
          calConfigs.classField = catValues[1];
          if (calConfigs.className)
            calConfigs.className = null;
        }
        else { //[0] assumed to be "Static"
          const categoryId = catValues[1]; //Will be the uniqueId, need to get the display name next
          if (this.props.categories) {
            this.props.categories.every((category:ICategoryItem) => {
              if (category.uniqueId === categoryId) {
                calConfigs.className = category.name; //store the display name instead
                if (calConfigs.classField)
                  calConfigs.classField = null;
                return false; //exit
              }
              else return true; //keep looping
            });
          }
        }
      }

      //Add Group/Row props to configs (groupField and groupId)
      //Split on the : char to determine if a field or category was selected
      if (calendar.group) {
        const catValues = calendar.group.split(":");
        if (catValues[0] === "Field") {
          calConfigs.groupField = catValues[1];
          if (calConfigs.groupId)
            calConfigs.groupId = null;
        }
        else { //[0] assumed to be "Static"
          calConfigs.groupId = catValues[1]; //Will be the uniqueId
          if (calConfigs.groupField)
            calConfigs.groupField = null;
        }
      }

      return calConfigs;
  }

  private getSPItemDates(list:IListItem, listConfigs:IListConfigs, itemData:any): IItemDateInfo {
    function getFieldValue(name:string) {
      if (name == null)
        return null;

      if (itemData.getAttribute) //SOAP
        return itemData.getAttribute("ows_" + name);
      else
        return itemData[name];
    }

    const allDayEvent = (getFieldValue("fAllDayEvent") === "1" || getFieldValue("fAllDayEvent") === true ? true : false);
    let strStartDateValue = getFieldValue(list.startDateField);
    if (allDayEvent && strStartDateValue)
      strStartDateValue = strStartDateValue.split("Z")[0]; //Drop the zulu designation to make it handle date as local

    let eventStartDate; //might be null for custom lists where the field isn't required/populated
    if (strStartDateValue) {
      //Check if date is valid first (calculated columns for example "datetime;#2024-03-31T22:00:00Z")
      const splits = strStartDateValue.split(";#");
      strStartDateValue = (splits[1] || splits[0]); //0 index is for other/"regular" fields
      eventStartDate = this.formatDateFromSOAP(strStartDateValue);
    }

    let eventEndDate; //might be null for custom lists where the field isn't required/populated
    if (list.endDateField) {
      let strEndDateValue = getFieldValue(list.endDateField);
      if (strEndDateValue) {
        if (allDayEvent)
          strEndDateValue = strEndDateValue.split("Z")[0]; //Drop the zulu designation to make it handle date as local

        //Check if date is valid first (calculated columns for example "datetime;#2024-03-31T22:00:00Z")
        const splits = strEndDateValue.split(";#");
        strEndDateValue = (splits[1] || splits[0]); //0 index is for other/"regular" fields
        eventEndDate = this.formatDateFromSOAP(strEndDateValue);
        //Check for non-calendar list dates in which no time is provided...(toLocaleTimeString() == "12:00:00 AM")
        if (list.isCalendar === false && listConfigs.extendEndTimeAllDay && eventEndDate.getHours() === 0 && eventEndDate.getMinutes() === 0) {
          //...change the time to the end of day
          eventEndDate.setDate(eventEndDate.getDate() + 1);
          eventEndDate.setSeconds(eventEndDate.getSeconds() - 1);
        }
      }
    }

    return {
      eventStartDate: eventStartDate,
      eventEndDate: eventEndDate
    }
  }

  private buildSPOItemObject(list:IListItem, listConfigs:IListConfigs, fieldKeys:string[], itemData:any, existingId?:number):any {
    //Helper function
    function getFieldValue(name:string) {
      if (name == null)
        return null;

      if (itemData.getAttribute) //SOAP
        return itemData.getAttribute("ows_" + name);
      else
        return itemData[name];
    }

    const itemDateInfo = this.getSPItemDates(list, listConfigs, itemData);
    if (itemDateInfo.eventStartDate == null) //must have a start date
      return {};

    //Get the "title" value
    let strTitle = "[No Title]"; //default value
    if (list.titleField) {
      strTitle = (getFieldValue(list.titleField) || strTitle);
      //Check for and handle calculated & lookup column formatting
      const splits = strTitle.split(";#"); //string;#Custom title
      strTitle = (splits[1] || splits[0]); //0 index is for other/"regular" fields
      strTitle = this.getFieldMappedValue(listConfigs, list.titleField, strTitle);
    }
    else { //Fallback
      strTitle = (getFieldValue("Title") || strTitle);
      strTitle = this.getFieldMappedValue(listConfigs, "Title", strTitle);
    }

    //Get item type
    let objType = getFieldValue("FSObjType") as string;
    if (objType)
      objType = objType.split(";#")[1]; //Value of "0" for item & "1" for folder
    //To check for DocSet: ows_ProgId='1/2/3;#Sharepoint.DocumentSet'

    //Build the event obj
    const oEvent = {
      id: (existingId || IdSvc.getNext()),
      spId: getFieldValue("ID"),
      objType: objType,
      encodedAbsUrl: getFieldValue("EncodedAbsUrl"),
      sourceObj: list,
      content: this.filterTextForXSS(strTitle),
      //title: elem.getAttribute("ows_Title"), //Tooltip
      start: itemDateInfo.eventStartDate,
      //end: elem.getAttribute("ows_EndDate"),
      type: "range", //Changed later as needed
      //className: //assigned next
      //group: list.groupId //assigned next
    } as any;
    
    //Add end date if applicable
    if (itemDateInfo.eventEndDate) {
      oEvent.end = itemDateInfo.eventEndDate;
      //Force single day events as point?
      //if (this.props.singleDayAsPoint && (elem.getAttribute(item.startFieldName).substring(0, 10) == elem.getAttribute(item.endFieldName).substring(0, 10)))
      //Better handling for user's time zone
      if (this.props.singleDayAsPoint && (itemDateInfo.eventStartDate.toLocaleDateString() === itemDateInfo.eventEndDate.toLocaleDateString()))
        oEvent.type = "point";
    }
    else //no end date, so make it a point
      oEvent.type = "point";

    //Add class/category
    if (listConfigs.className)
      oEvent.className = this.props.ensureValidClassName(listConfigs.className);
    //Apply class by category field
    else if (listConfigs.classField) {
      const classFieldValue = getFieldValue(listConfigs.classField); //calConfigs.classField
      if (classFieldValue) {
        const categorySplit = this.handleMultipleSOAPValues(classFieldValue);
        //NOTE: "regular" text values without ;# still result in ["Single value"] array, so below logic still works
        let mappedValue = null;
        if (listConfigs.multipleCategories == 'useLast')
          mappedValue = this.getFieldMappedValue(listConfigs, listConfigs.classField, categorySplit[categorySplit.length-1]);
        else //default to use first value
          mappedValue = this.getFieldMappedValue(listConfigs, listConfigs.classField, categorySplit[0]);
        
        //Set it
        oEvent.className = this.props.ensureValidClassName(mappedValue);
      }
    }

    if (oEvent.className && oEvent.className === this.props.ensureValidClassName(this.props.holidayCategories)) {
      oEvent.type = "background"; //change to background
      oEvent.group = null; //apply to entire timeline
    }
    
    //Add data to the event object (for later tooltip template processing)
    fieldKeys.forEach(field => {
      //Skip these fields to prevent their above defined value from being overwritten
      if (field === "id" || field === "content" || field === "start" || field === "end" || field === "type" || field === "className")
        return;

      let fieldValue = getFieldValue(field);
      if (fieldValue) {
        oEvent[field] = (fieldValue || ""); //save initial value
        //Look for special fields
        if (field == "Description") {
          //Remove blanks
          if (fieldValue === "<div></div>" || fieldValue === "<div></div><p>​</p>") //last one has a *hidden* character
            fieldValue = "";
          //If HTML text (versus plain text), starts with < character
          if (fieldValue.indexOf("<") === 0) {
            //Remove break at end of paragraph and empty paragraphs and paragraph with a *hidden* &ZeroWidthSpace; character
            fieldValue = fieldValue.replace(/<br><\/p>/g, "</p>").replace(/<p><\/p>/g, "").replace(/<p>​<\/p>/g, ""); //last one has a *hidden* character
            /*//Wrap just to make "finding" easier in next steps
            var $desc = $("<div>" + fieldValue + "</div>");
            //Remove padding from first paragraph
            $desc.find("p:first").css("margin-top", "0");
            //Remove white background styling
            $desc.find("*").each(function() {
                if ($(this).css("background-color") == "#ffffff" || $(this).css("background-color") == "rgb(255, 255, 255)")
                  $(this).css("background-color", "inherit");
            })
            //$("table").css("width", "100%");
            
            //Extract the new HTML
            fieldValue = $desc[0].outerHTML;
            */
            const divWrapper = document.createElement("div");
            divWrapper.innerHTML = fieldValue;
            const firstP = divWrapper.querySelector("p:first-child") as HTMLElement;
            //Remove padding from first paragraph
            if (firstP)
              firstP.style.marginTop = "0px";
            divWrapper.querySelectorAll("*").forEach((elem: HTMLElement) => {
              //Remove white and gray background colors
              if (elem.style.backgroundColor === "#ffffff" || elem.style.backgroundColor === "rgb(255, 255, 255)" ||
                    elem.style.backgroundColor === "#dfdfdf")
                elem.style.backgroundColor = "inherit";
            });
            fieldValue = divWrapper.outerHTML;
          }
          else { //Plain text
            fieldValue = fieldValue.replace(/\r?\n/g, "<br>");
            //Make sure it's not ending with a line break
            fieldValue = fieldValue.replace(/<br>$/, "");
          }
          //Set prop with updated value
          oEvent[field] = this.filterTextForXSS(fieldValue);
        }
        else {
          //Handle Number/Currency field values like 1.00000000000000 & 10.2000000000000
          //@ts-ignore @typescript-eslint/TS2550 (for endsWith)
          if (isNaN(Number(fieldValue)) === false && fieldValue.toString().endsWith("0000")) {
            //Number() removes the extra 0s
            fieldValue = Number(fieldValue).toString();
          }

          //Check for lookup (and choice?) fields with "123;#Display Name" format and multiple value fields also using ";#" delimeter
          // let fieldSplit = fieldValue.split(";#");
          // if (fieldSplit.length == 2)
          //   oEvent[field] = fieldSplit[1];
          // else if (fieldSplit.length > 2) {
          //   //Multiple value field (";#Chevy;#Porsche;#"); remove potential blank entries in split
          //   fieldSplit = fieldSplit.filter(i => {return i});
          //   oEvent[field] = fieldSplit.join(", ");
          // }
          const fieldSplit = this.handleMultipleSOAPValues(fieldValue);
          oEvent[field] = this.filterTextForXSS(fieldSplit.join(", "))
          oEvent[field] = this.getFieldMappedValue(listConfigs, field, oEvent[field]);
        }
      }
    });

    return oEvent;
  }

  private getFieldMappedValue(configsObj:any, field:string, fieldValue:string, forTooltip?:boolean): any {
    //Handle any user provided mappings
    /* Example format
    {
      "showAs": { --> This is the valueMappingObj
        "oof": "Out of Office",
        "free": "Free"
      },
      "nextField": {...}
    }
    */
    if (configsObj.fieldValueMappings) {
      const valueMappingObj = configsObj.fieldValueMappings[field]
      if (valueMappingObj && typeof valueMappingObj === "object") {
        //Get the new [mapped] value
        const newValue = valueMappingObj[fieldValue];
        
        //Return original value in tooltip if property specified
        if (forTooltip && valueMappingObj._shownInTooltip === false)
          return fieldValue;

        return (newValue || fieldValue);
      }
    }
    return fieldValue;
  }

  //For calEvent prop: Extend Graph.Event interface with implicitly defined index signature (to support eventObj["keyFieldName"] retrieval)
  private buildCalendarEventObject(calendar:ICalendarItem, calConfigs:ICalendarConfigs, fieldKeys:string[], calEvent:MicrosoftGraph.Event & {[key: string]:any}, existingId?:number):any {
    //Get the "title" value
    let strTitle = "[No Title]"; //default value, changed next
    if (calEvent.subject != null && calEvent.subject.trim() != "") {
      strTitle = calEvent.subject;
    }
    strTitle = this.getFieldMappedValue(calConfigs, "subject", strTitle);
    
    //Process "non-dates" ("0001-01-01T00:00:00Z" is returned for private events)
    if (calEvent.createdDateTime === "0001-01-01T00:00:00Z") { //had: calEvent.sensitivity == "private" && 
      calEvent.createdDateTime = null;
      calEvent.lastModifiedDateTime = null;
    }

    /*{ //Outlook dates are returned in UTC but without a trailing "Z" character
        dateTime: "2024-03-22T00:00:00.0000000" or "2024-03-18T12:30:00.0000000"
        timeZone: "UTC"
      }*/
    //All day event *end* dates show as the *next* day with 00:00:00 time
    let eventEndDate;
    if (calEvent.isAllDay) { //Change the end to be the end of the correct day
      eventEndDate = new Date(calEvent.end.dateTime);
      eventEndDate.setMinutes(eventEndDate.getMinutes() -1); //shows as 23:59:00
    }
    else
      eventEndDate = new Date(calEvent.end.dateTime + "Z");

    //Build the event obj
    const oEvent = {
      id: IdSvc.getNext(),
      //spId: TODO: Rename to sourceId? for calEvent.id
      eventId: calEvent.id,
      //encodedAbsUrl: for SPO,
      calEventWebLink: calEvent.webLink,
      sourceObj: calendar,
      content: this.filterTextForXSS(strTitle),
      //title: elem.getAttribute("ows_Title"), //Tooltip
      start: new Date(calEvent.start.dateTime + (calEvent.isAllDay ? "" : "Z")), //All day events treated as local time
      end: eventEndDate,
      type: "range", //Changed later as needed
      //className: //assigned next
      //group: list.groupId //assigned next
    } as any;
    
    //Force single day events as point?
    if (this.props.singleDayAsPoint && (oEvent.start.toLocaleDateString() === oEvent.end.toLocaleDateString()))
      oEvent.type = "point";

    //Add class/category
    if (calConfigs.className)
      oEvent.className = this.props.ensureValidClassName(calConfigs.className);
    //Apply class by category field
    else if (calConfigs.classField) {
      //Check for an array
      const classFieldValue = calEvent[calConfigs.classField];
      if (Array.isArray(classFieldValue)) {
        let mappedValue = null;
        if (calConfigs.multipleCategories === 'useLast')
          mappedValue = this.getFieldMappedValue(calConfigs, calConfigs.classField, classFieldValue[classFieldValue.length-1]);
        else //default to use first value
          mappedValue = this.getFieldMappedValue(calConfigs, calConfigs.classField, classFieldValue[0]);
        
        //Set it
        oEvent.className = this.props.ensureValidClassName(mappedValue);
      }
      else {
        const mappedValue = this.getFieldMappedValue(calConfigs, calConfigs.classField, classFieldValue);
        oEvent.className = this.props.ensureValidClassName(mappedValue);
      }
    }

    if (oEvent.className && oEvent.className === this.props.ensureValidClassName(this.props.holidayCategories)) {
      oEvent.type = "background"; //change to background
      oEvent.group = null; //apply to entire timeline
    }
    
    // //Call custom function if provided
    // if (TC.settings.beforeEventAdded)
    //   oEvent = TC.settings.beforeEventAdded(oEvent, $(this), cal);

    //Add data to the event object (for later tooltip template processing)
    fieldKeys.forEach(field => {
      //Skip these fields to prevent their above defined value from being overwritten
      if (field === "id" || field === "content" || field === "start" || field === "end" || field === "type" || field === "className")
        return;

      let fieldValue = calEvent[field]; //TODO: support object values like organizer.emailAddress.address
      //let wasMapped = true;
      //Map certain field names to help match to existing SP calendar fields
      switch (field) {
        //case "location":
        case "Location":
          fieldValue = calEvent.location.displayName;
          fieldValue = this.getFieldMappedValue(calConfigs, "location", fieldValue, true);
          oEvent.Location = this.filterTextForXSS(fieldValue);
          break;

        //case "categories":
        case "Category":
          fieldValue = calEvent.categories.join(", ");
          fieldValue = this.getFieldMappedValue(calConfigs, "categories", fieldValue, true);
          oEvent.Category = this.filterTextForXSS(fieldValue);
          break;

        //case "body":
        case "Description":
          fieldValue = (calEvent.body && calEvent.body.content || "");
          fieldValue = this.getFieldMappedValue(calConfigs, "body", fieldValue, true);
          oEvent.Description = this.filterTextForXSS(fieldValue);
          break;

        //case "organizer":
        case "Author":
          fieldValue = (calEvent.organizer && calEvent.organizer.emailAddress.name || "");
          fieldValue = this.getFieldMappedValue(calConfigs, "organizer", fieldValue, true);
          oEvent.Author = this.filterTextForXSS(fieldValue);
          break;

        // case "Editor":
        //   fieldValue = calEvent.??; //perhaps an extended MAPI property
        //   oEvent["Editor"] = this.getFieldMappedValue(calConfigs, field, fieldValue);
        //   break;

        //case "createdDateTime":
        case "Created":
          fieldValue = calEvent.createdDateTime;
          fieldValue = this.getFieldMappedValue(calConfigs, "createdDateTime", fieldValue, true);
          oEvent.Created = this.filterTextForXSS(fieldValue);
          break;
        
        //case "lastModifiedDateTime":
        case "Modified":
          fieldValue = calEvent.lastModifiedDateTime;
          fieldValue = this.getFieldMappedValue(calConfigs, "lastModifiedDateTime", fieldValue, true);
          oEvent.Modified = this.filterTextForXSS(fieldValue);
          break;

        case "charmIcon":
          fieldValue = (calEvent.singleValueExtendedProperties && calEvent.singleValueExtendedProperties[0] &&
                          calEvent.singleValueExtendedProperties[0].value || "");
          if (fieldValue === "None")
            fieldValue = ""; //set to blank instead

          fieldValue = this.getFieldMappedValue(calConfigs, "charmIcon", fieldValue, true);
          oEvent.charmIcon = this.filterTextForXSS(fieldValue);
          break;

        default:
          fieldValue = (this.getFieldMappedValue(calConfigs, field, fieldValue, true) || "");
          oEvent[field] = this.filterTextForXSS(fieldValue);
      }
    });

    return oEvent;
  }

  private getSharePointEvents(): Promise<any[]> {
    return Promise.all(this.props.lists.map((list:IListItem) => {
      const configs = this.buildListConfigs(list);
      //Only process valid entries
      if (configs.visible === false)
        return; //skip this one
      else
        return this.queryList(list, configs);
    }))
  }

  private calendarEventsSoapEnvelope(list:IListItem, listConfigs:IListConfigs, nextPageDetail?:string):string {
    const fieldKeys = this.getFieldKeys();
    const includeRecurrence = (list.isCalendar !== false);
  
    let returnVal = "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/'><soapenv:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" + 
			"<listName>" + list.list + "</listName>" + 
			"<viewFields><ViewFields>" + 
				//"<FieldRef Name='Title' />" + 
        (list.titleField ? "<FieldRef Name='" + list.titleField + "' />" : "") +
				//"<FieldRef Name='Location' />" + 
				"<FieldRef Name='EventDate' />" +
				"<FieldRef Name='EndDate' />" + 
				(includeRecurrence ? "<FieldRef Name='fRecurrence' />" : "") + 
				(includeRecurrence ? "<FieldRef Name='RecurrenceData' />" : "") + 
				"<FieldRef Name='EncodedAbsUrl' />" +
        //fAllDayEvent seems to be included by default
        //"<FieldRef Name='Author' />" + //now pulled from tooltip template if included there (see fieldKeys loop below)
        //"<FieldRef Name='Editor' />" +
        (listConfigs.classField ? "<FieldRef Name='" + listConfigs.classField + "' />" : "") +
        (listConfigs.groupField ? "<FieldRef Name='" + listConfigs.groupField + "' />" : "");

    //Add fields used in the tooltip
    fieldKeys.forEach(field => {
      returnVal += "<FieldRef Name='" + field + "' />";
    });
    
    //Add date fields
    if (list.startDateField)
      returnVal += "<FieldRef Name='" + list.startDateField + "' />";
    if (list.endDateField)
      returnVal += "<FieldRef Name='" + list.endDateField + "' />";
    
    returnVal += "</ViewFields></viewFields>" + 
    //Set a rowLimit, without this the default is only 30 items
    "<rowLimit Paged=\"TRUE\">1000</rowLimit>"; //Even when setting this to 5000 SP only returns...
    //  999 events for calendars
    //  1000 events for non-calendars/lists
    //...and requires you to use pagination for next batch of events.
    
    let query = "<query><Query><Where>"; //<Where> is the start of the ViewQuery property from calendar views

    //if (cal.filter != "") //this was needed to prevent the general filter from applying when no calendar filter is wanted
    //	query += (cal.filter ? "<And>" + cal.filter : (TC.settings.filter ? "<And>" + TC.settings.filter : ""));

    //Determine applicable filter
    let filterToUse = null;
    if (list.viewFilter)
      filterToUse = list.viewFilter;
    else if (listConfigs.camlFilter && listConfigs.camlFilter.trim() !== "")
      filterToUse = listConfigs.camlFilter;

    if (filterToUse)
      query += "<And>" + filterToUse;
    
    //Handle custom lists/non-calendars
    if (list.isCalendar === false) {
      //Was an end date field provided?
      if (list.endDateField) //use both dates from the non-calendar
      query += "<And><Geq><FieldRef Name='" + list.endDateField + "'/><Value Type='DateTime'>" + this.getMinDate().toISOString() + "</Value></Geq>" +
        "<Leq><FieldRef Name='" + list.startDateField + "'/><Value Type='DateTime'>" + this.getMaxDate().toISOString() + "</Value></Leq>" +
      "</And>";
      else //Only a start date was given
        query += "<Geq><FieldRef Name='" + list.startDateField + "'/><Value Type='DateTime'>" + this.getMinDate().toISOString() + "</Value></Geq>";
    }
    //Handle classic SP calendars
    else {
      //Standard calendar events filter
      if (includeRecurrence) {
        query +=
        "<And>" +
          //Prevents returning older events that aren't in viewable range
          //Moving this below <DateRangeOverlap> caused threshold errors in large calendars (and also not having it)
          "<Geq><FieldRef Name='EndDate'/><Value Type='DateTime'>" + this.getMinDate().toISOString() + "</Value></Geq>" +
          "<DateRangesOverlap>" + 
            "<FieldRef Name='EventDate' />" +
            "<FieldRef Name='EndDate' />" + 
            "<FieldRef Name='RecurrenceID' />" + 
            "<Value Type='DateTime'><Year/></Value>" + //No value or <Year/> are the same
          "</DateRangesOverlap>" +
        "</And>";
      }
    }
    
    if (filterToUse)
      query += "</And>";
    
    query += "</Where>";

    
    returnVal += query + /*"<OrderBy>" + 
      "<FieldRef Name='EventDate' />" + 
    "</OrderBy>" +*/
    "</Query></query><queryOptions><QueryOptions>" + 
      //(calendarDate ? "<CalendarDate>" + calendarDate + "</CalendarDate>" : "") + //today is the default if no element provided
      (nextPageDetail ? "<Paging ListItemCollectionPositionNext='" + nextPageDetail + "' />" : "" ) + 
      (includeRecurrence ? "<RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion>" : "") + 
      (includeRecurrence ? "<ExpandRecurrence>TRUE</ExpandRecurrence>" : "") + 
      (includeRecurrence ? "<RecurrenceOrderBy>TRUE</RecurrenceOrderBy>" : "") +
      "<ViewAttributes Scope='RecursiveAll' />" +
      //"<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" +
      "<ViewFieldsOnly>TRUE</ViewFieldsOnly>" +
      (listConfigs.dateInUtc === false ? "" : "<DateInUtc>TRUE</DateInUtc>") + //True returns dates as "2023-10-10T06:00:00Z" versus "2023-10-10 08:00:00"
    "</QueryOptions></queryOptions></GetListItems></soapenv:Body></soapenv:Envelope>";
    
    return returnVal;
  }

  private handleMultipleSOAPValues(fieldValue:string): string[] {
    //Check for multiple value fields (single value fields work fine and result in ["Single value"] array)
    let fieldSplit = fieldValue.split(";#");

    /* NOTE: Field values can have special formatting for some fields...
    Calculated field:
    string;#Custom title

    Multi-Choice field:
    ;#Division Event;# (when a single value selected)
    ;#Division Event;#Family Event;# (when multiple values selected)

    Lookup & Person & Managed Metadata fields:
    1;#Some test item
    9;#VASILOFF, MICHAEL D CTR USAF USAFE USAFE CS/CSK

    Multi-Lookup & Multi-Person & Multi-Managed Metadata fields
    9;#VASILOFF, MICHAEL D CTR USAF USAFE USAFE CS/CSK
    8;#SharePoint Online;#9;#OneDrive
    27;#HILLMAN, MATTHEW G CTR USAF USAFE USAFE-AFAFRICA/CSK;#9;#VASILOFF, MICHAEL D CTR USAF USAFE USAFE CS/CSK

    "Regular" values without ;# result in ["Single value"] array
    */

    //Look for Calculated fields first (since the value left of ;# isn't a "real" value within the .split)
    if (fieldSplit.length === 2 && isNaN(Number(fieldSplit[0]))) {
      //Remove the first value from the array (leaving a single entry array)
      fieldSplit.shift();
    }
    
    //Check for multiple-value fields 
    //Remove potential blank entries in split (from multi-choice fields)
    fieldSplit = fieldSplit.filter(i => {return i});

    //A single value was found (either a "regular" value or from a Calculated field or a multi-choice field)
    if (fieldSplit.length === 1) {
      //groupFieldValue = groupSplit[0];
      return fieldSplit;
    }
    
    //Check for *single* value lookup/person/managed metadata field (has a number in the 0 index)
    // if (groupSplit.length == 2 && !isNaN(Number(groupSplit[0]))) {
    //   groupFieldValue = groupSplit[1];
    // }
    // else {
    {
      //More than one, *real* value is in the field
      //Look for Multi-Lookup & Multi-Person & Multi-Managed Metadata fields where the even index is an integer
      let onlyIntFound = true;
      for (let i=0; i < fieldSplit.length; i+=2) {
        //See if value is *not* an integer
        if (isNaN(Number(fieldSplit[i])) && parseInt(fieldSplit[i]).toString() != fieldSplit[i]) {
          onlyIntFound = false;
          break;
        }
      }
      //If only integers were found (in the even indexes)
      if (onlyIntFound) {
        //Remove the even index items by returning only the odds
        fieldSplit = fieldSplit.filter((value:string, index:number) => {
          return index % 2 !== 0; //Index is odd
        });
      }
    }

    return fieldSplit;
  }

  private async queryList(list:IListItem, listConfigs:IListConfigs, nextPageDetail?:string): Promise<void> {
    return await this.props.context.spHttpClient.post(list.siteUrl + "/_vti_bin/lists.asmx", SPHttpClient.configurations.v1,
    {
      headers: [
        ["Accept", "application/xml, text/xml, */*; q=0.01"],
        ["Content-Type", 'text/xml; charset="UTF-8"']
      ],
      body: this.calendarEventsSoapEnvelope(list, listConfigs, nextPageDetail)
    }).then((response: SPHttpClientResponse) => response.text())
    .then((strXml: any) => {
      // //Check for problems such as access denied to the site/web object
      // //They won't have an rs:data element with ItemCount attribute
      // if ($(jqxhr.responseXML).SPFilterNode("rs:data").attr("ItemCount") == null) {
      //   var strError = cal.listName + " returned invalid response";
      //     var isWarning = false;
      //     var isError = false;
      //     if (jqxhr.responseXML) {
      //         var msg = ($("title", jqxhr.responseXML).text() || "").trim();
      //         if (msg == "Access required") {
      //           isWarning = true;
      //           strError += ": Could not access site: " + cal.siteUrl;
      //         }
      //         else {
      //           strError += ": " + ($("h1.ms-core-pageTitle", jqxhr.responseXML).text() || "").trim();
      //           isError = true;
      //         }
      //     }
      //     else
      //       isError = true;
      //     TC.log(strError);
      //     dispatcher.queryCompleted(cal, isWarning, isError);
          
      //   return; //don't proceed with the below
      // }
      //TODO: Or errors like this
      // <soap:Body>
      //   <soap:Fault>
      //     <faultcode>soap:Server</faultcode>
      //     <faultstring>Exception of type 'Microsoft.SharePoint.SoapServer.SoapServerException' was thrown.</faultstring>
      //     <detail>
      //       <errorstring xmlns="http://schemas.microsoft.com/sharepoint/soap/">The attempted operation is prohibited because it exceeds the list view threshold.</errorstring>
      //       <errorcode xmlns="http://schemas.microsoft.com/sharepoint/soap/">0x80070024</errorcode>
      //     </detail>
      //   </soap:Fault>
      // </soap:Body>

      //Valid data response
      // <soap:Body>
      //     <GetListItemsResponse xmlns="...">
      //       <GetListItemsResult>
      //         <listitems xmlns:s='...' ... xmlns:z='#RowsetSchema'>
      //           <rs:data ItemCount="552" ListItemCollectionPositionNext="Paged=Next&amp;p_StartTimeUTC=20261202T130001Z">
      //             <z:row ows_Title='Event 495'...
      
      //At this point we should have a valid list response
      //let numOfValidItems = 0;
      let lastStartDate:Date = null;
      const fieldKeys = this.getFieldKeys();
      
      let pagingDetails:string = null;
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(strXml, "application/xml");
      xmlDoc.querySelectorAll("*").forEach(elem => {
        if (elem.nodeName === "rs:data")
          //Store this for use after the forEach
          pagingDetails = elem.getAttribute("ListItemCollectionPositionNext"); //ItemCount is another
        
        //Loop over the event/data results
        else if(elem.nodeName === "z:row") { //actual data is here
          const itemDateInfo = this.getSPItemDates(list, listConfigs, elem);
          if (itemDateInfo.eventStartDate == null) //Cannot add events with no start date
              return; //skip this one

          lastStartDate = itemDateInfo.eventStartDate; //saved for later
          
          //CAML returns *recurring* events for a whole year prior to and after today regardless of any EventDate/EndDate filters
          //So only include events that fall within the requested date range
          if ((itemDateInfo.eventEndDate == null || itemDateInfo.eventEndDate >= this.getMinDate()) && itemDateInfo.eventStartDate <= this.getMaxDate()) {
            //numOfValidItems++;
            
            const oEvent = this.buildSPOItemObject(list, listConfigs, fieldKeys, elem);
            
            //Add group (row/swimlane)
            let multipleValuesFound = false;
            if (listConfigs.groupId) {
              //When a class field (Category) is selected and a specific Row is selected
              if (listConfigs.classField && listConfigs.limitHolidayToRow === false 
                    && oEvent.className && oEvent.className === this.props.ensureValidClassName(this.props.holidayCategories)) {
                oEvent.type = "background"; //change to background
                oEvent.group = null; //apply to entire timeline
              }
              else
                oEvent.group = listConfigs.groupId;
            }
            else if (listConfigs.groupField && this.props.groups) {
              //Find the associated group to assign the item to
              let groupFieldValue = elem.getAttribute("ows_" + listConfigs.groupField);
              if (groupFieldValue) {
                const groupSplit = this.handleMultipleSOAPValues(groupFieldValue);

                //Look for "regular" values without ;# (they result in ["Single value"] array)
                if (groupSplit.length === 1) {
                  groupFieldValue = groupSplit[0];
                  groupFieldValue = this.getFieldMappedValue(listConfigs, listConfigs.groupField, groupFieldValue);
                }
                else {
                  //More than one, *real* value is in the field
                  multipleValuesFound = true;

                  //Create a duplicate event for each selected group value
                  groupSplit.forEach(groupName => {
                    const eventClone = structuredClone(oEvent);
                    //Above duplicates the event object
                    eventClone.id = IdSvc.getNext(); //Set a new I
                    //Map value if applicable
                    groupName = this.getFieldMappedValue(listConfigs, listConfigs.groupField, groupName);

                    //Find the associated group from it's name
                    this.props.groups.every((group:IGroupItem) => {
                      if (group.name === groupName) {
                        eventClone.group = group.uniqueId;
                        return false; //exit
                      }
                      else return true; //keep looping
                    });

                    //Add the clone to the DataSet
                    this._dsItems.add(eventClone);
                  });
                }

                //Finalize single value events
                if (multipleValuesFound === false) {
                  //Find the associated group from it's name
                  this.props.groups.every((group:IGroupItem) => {
                    if (group.name === groupFieldValue) {
                      //When a class field (Category) is selected and a *field* for Row selected
                      if (listConfigs.classField && listConfigs.limitHolidayToRow === false 
                        && oEvent.className && oEvent.className === this.props.ensureValidClassName(this.props.holidayCategories)) {
                        oEvent.type = "background"; //change to background
                        oEvent.group = null; //apply to entire timeline
                      }
                      else
                        oEvent.group = group.uniqueId;
                      return false; //exit
                    }
                    else return true; //keep looping
                  });
                }
              } //There is a groupFieldValue
            } //A groupField was selected && there are this.props.groups

            //Add event/item to the DataSet
            if (multipleValuesFound === false)
              this._dsItems.add(oEvent);
          }
        }
      }); //end SOAP response forEach

      //Check if more data should be queried
      if (pagingDetails) {
        let validStartTime = true; //initial
        //Calendars return these formats (even if user specifies dateInUtc=false):
        // > Paged=Next&p_StartTimeUTC=20250702T120001Z - Valid result!
        // > Paged=Next&p_StartTimeUTC=18991230T000001Z - Invalid date (starts with 1899)
        // > Paged=Next&p_StartTimeUTC=00000101T588376552265528Z - Invalid, too long
        //    Using these invalid results in the next call it returns the same items, so prevent these follow-on calls
        //Non-calendars return this format:
        // > Paged=TRUE&p_ID=40
        const searchParams = new URLSearchParams(pagingDetails);
        if (list.isCalendar) {
          const strTemp = searchParams.get("p_StartTimeUTC");
          //@ts-ignore @typescript-eslint/TS2550 (for startsWith)
          if (strTemp && (strTemp.indexOf("T") !== 8 || strTemp.startsWith("1899"))) {
            validStartTime = false;
            console.log(list.listName + " returned an invalid ListItemCollectionPositionNext: " + pagingDetails + " (ignoring)");
          }
        }
        else { //non-calendar list
          const strTemp = searchParams.get("p_ID");
          if (!strTemp)
              validStartTime = false;
        }

        if (validStartTime && lastStartDate < this.getMaxDate()) {						
          //Need to keep "&" character encoded for next SOAP call
          pagingDetails = pagingDetails.replace("&p_", "&amp;p_");
          //Query for more events
          return this.queryList(list, listConfigs, pagingDetails);
        }
      }
    })
    .catch((error: any) => {
      console.error(error);
    });
  }

  private getOutlookEvents(): Promise<any[]> {
    return Promise.all(this.props.calsAndPlans.map((calendar:ICalendarItem) => {
      const configs = this.buildCalendarConfigs(calendar);
      
      //Only process valid entries
      if (configs.visible == false || calendar.persona == null || calendar.persona.length == 0)
        return; //skip this one
      else
        return this.queryCalendar(calendar, configs, null, 0);
    }))
  }

  private async queryCalendar(calObj:ICalendarItem, calConfigs:ICalendarConfigs, existingEvent:any, skipNumber:number, startDate?:string, endDate?:string): Promise<void> {
    //First check for Flow3 to prevent full page redirection due to Conditional Access policy block on token issuance
    const params = new URLSearchParams(window.location.search);
    if (params.has("ignoreDownloadsCheck") || !this.props.context.pageContext.legacyPageContext.blockDownloadsExperienceEnabled)
    //First Promise return but later return again from the Graph call
    return await this.props.graphClient.then((client:MSGraphClientV3): void => {
      const appendValidGraphEventProp = (name:string): string => {
        //Only certain Event properties can be allowed in Graph query or an error is received:
        //HTTP 400: Could not find a property named 'Category' on type 'Microsoft.OutlookServices.Event'.
        switch(name) {
          //In addition to those defined in selectFields above...
          case "originalStartTimeZone":
          case "originalEndTimeZone":
          case "iCalUId":
          case "reminderMinutesBeforeStart":
          case "isReminderOn":
          case "hasAttachments":
          case "bodyPreview":
          case "importance":
          case "sensitivity":
          case "isCancelled":
          case "isOrganizer":
          case "responseRequested":
          case "showAs":
          case "type":
          case "onlineMeetingUrl":
          case "onlineMeeting":
          case "isOnlineMeeting":
          case "onlineMeetingProvider":
          case "allowNewTimeProposals":
          case "hideAttendees":
            return "," + name;
        }
        return null;
      };
      const fieldKeys = this.getFieldKeys();
      
      //Build initial list of fields to select (will be augmented by tooltip fieldKeys)
      //singleValueExtendedProperties doesn't need to be selected (the $expand seems to include it already)
      let selectFields = "id,organizer,createdDateTime,lastModifiedDateTime,categories,subject,body,start,end,location,isAllDay,webLink";
      //Duplicates are OK (if categories (above) is classField and/or showAs is groupField and is part of fieldKeys)
      selectFields += (appendValidGraphEventProp(calConfigs.classField) || "");
      selectFields += (appendValidGraphEventProp(calConfigs.groupField) || "");

      //Add fields used in the tooltip
      fieldKeys.forEach(field => {
        selectFields += (appendValidGraphEventProp(field) || "");
      });

      //Build API URL based on user or group calendar
      let existingEventId: string = null;
      if (existingEvent && existingEvent.eventId)
        existingEventId = existingEvent.eventId;
      let apiURL = "";
      const resourceId = calObj.resource.split(":")[1]; //format "calendar:Id"
      if (calObj.persona[0].personaType === "user")
        apiURL = "/users/" + calObj.persona[0].mail + "/calendars/" + resourceId + 
          //Single event query (for post user clicking an item) or multiple events query
          (existingEventId ? "/events/" + existingEventId : "/calendarView");
      else //assumed to be a group
        apiURL = "/groups/" + calObj.persona[0].key + 
          //Single event query (for post user clicking an item) or multiple events query
          (existingEventId ? "/events/" + existingEventId : "/calendarView");

      //Build API call variables
      if (!existingEventId && !startDate) {
        //Graph API limited to querying max 5 year date range/span
        if ((this.getMaxDate().getTime() - this.getMinDate().getTime()) / (1000*60*60*24) > 1825) {
          startDate = this.getMinDate().toISOString();
          const tempDate = new Date(this.getMinDate().getTime());
          tempDate.setFullYear(tempDate.getFullYear()+3); //add 3 years
          endDate = tempDate.toISOString();
        }
        else {
          startDate = this.getMinDate().toISOString();
          endDate = this.getMaxDate().toISOString();
        }
      }
      const basicQueryStringParams = (existingEventId ? '' : `startDateTime=${startDate}&endDateTime=${endDate}`);
      const filter = (existingEventId ? "" : (calObj.filter ? calObj.filter.trim() : ""));
      
      //Get calender view events
      //@ts-ignore ("return" does work here)
      return client.api(apiURL).query(basicQueryStringParams)
      //TODO: 500 or higher?
      .select(selectFields).top(500).skip(skipNumber)
      .filter(filter)
      //Make the charm/icon value be included
      .expand("singleValueExtendedProperties($filter=id eq 'Integer {11000E07-B51B-40D6-AF21-CAA85EDAB1D0} Id 0x0027')")
      //Headers -> Prefer: outlook.timezone (string) -> If not specified dates are returned in UTC
      .get((error:GraphError, response:any, rawResponse?:any) => {
        if (error) {
          //console.log(error.message);
          //Could be this in case of missing Graph scopes or user permission to the resource
          //"code": "ErrorAccessDenied",
          //"message": "Access is denied. Check credentials and try again."
          //"statusCode": 403
        }
        else {
          let events:(MicrosoftGraph.Event & {[key: string]:any})[] = null;
          if (response.id) //single event
            events = [response];
          else //multiple events query
            events = response.value;

          events.forEach(calEvent => {
            const oEvent = this.buildCalendarEventObject(calObj, calConfigs, fieldKeys, calEvent);
            if (existingEvent)
            oEvent.id = existingEvent.id; //must set back to original id to ensure it is updated

            //Add group (row/swimlane)
            let multipleValuesFound = false;
            if (calConfigs.groupId) {
              oEvent.group = calConfigs.groupId;
              if (existingEventId) //need to update this existing item
                this._dsItems.update(oEvent);
            }
            else if (calConfigs.groupField && this.props.groups) {
              //Get value of group field
              let groupFieldValue = null as string | string[];
              switch (calConfigs.groupField) {
                case "categories":
                  groupFieldValue = calEvent.categories;
                  break;

                case "showAs":
                  //MicrosoftGraph.FreeBusyStatus = "unknown" | "free" | "tentative" | "busy" | "oof" | "workingElsewhere";
                  //Mapping option: "Unknown", "Free", "Tentative", "Busy", "Away"/Out of Office, "Working elsewhere"/Working Elsewhere
                  groupFieldValue = calEvent.showAs.toString();
                  break;

                case "charmIcon":
                  groupFieldValue = (calEvent.singleValueExtendedProperties && calEvent.singleValueExtendedProperties[0] &&
                    calEvent.singleValueExtendedProperties[0].value || "");
                  if (groupFieldValue === "None")
                    groupFieldValue = ""; //overwrite value
                  break;

                default:
                  groupFieldValue = calEvent[calConfigs.groupField];
              }

              //Find the associated group to assign the item to
              if (groupFieldValue) {
                if (Array.isArray(groupFieldValue)) {
                  multipleValuesFound = true;

                  if (existingEventId) {
                    //Now need to assume that there *could* have been multiple selections previously
                    //but the item was updated to only have one (we need to remove those prior copies)
                    //Find duplicate events and remove them
                    const itemEvents = this._dsItems.get({
                      filter: function (item:any) {
                        return (item.eventId === oEvent.eventId); // && item.id != oEvent.id
                      }
                    });
                    this._dsItems.remove(itemEvents);
                  }

                  //Create a duplicate event for each selected group value
                  groupFieldValue.forEach(groupName => {
                    const eventClone = structuredClone(oEvent);
                    //Above duplicates the event object
                    eventClone.id = IdSvc.getNext(); //Set a new ID
                    //Map value if applicable
                    groupName = this.getFieldMappedValue(calConfigs, calConfigs.groupField, groupName);

                    //Find the associated group from it's name
                    this.props.groups.every((group:IGroupItem) => {
                      if (group.name === groupName) {
                        eventClone.group = group.uniqueId;
                        return false; //exit
                      }
                      else return true; //keep looping
                    });

                    //Add the clone to the DataSet
                    this._dsItems.add(eventClone);
                  });
                }

                //Finalize single value events
                if (existingEventId == null && multipleValuesFound === false) {
                  //Map value if applicable
                  groupFieldValue = this.getFieldMappedValue(calConfigs, calConfigs.groupField, (groupFieldValue as string));
                  //Find the associated group from it's name
                  this.props.groups.every((group:IGroupItem) => {
                    if (group.name === groupFieldValue) {
                      oEvent.group = group.uniqueId;
                      return false; //exit
                    }
                    else return true; //keep looping
                  });
                }
              } //There is a groupFieldValue
            } //A groupField was selected && there are this.props.groups

            //Add event/item to the DataSet
            if (existingEventId == null && multipleValuesFound === false)
              this._dsItems.add(oEvent);
          });

          if (!existingEventId) {
            //Check if more data (paging) should be queried
            const nextLink = response["@odata.nextLink"] as string;
            if (nextLink) {
              //Query for more events (get the next page)
              const url = new URL(nextLink);
              skipNumber = Number(url.searchParams.get("$skip"));
              return this.queryCalendar(calObj, calConfigs, existingEvent, skipNumber, startDate, endDate);
            }
            //Check if another date range should be queried
            else if (new Date(endDate) < this.getMaxDate()) {
              const tempDate = new Date(endDate);
              tempDate.setMilliseconds(tempDate.getMilliseconds() + 1);
              startDate = tempDate.toISOString();
              tempDate.setFullYear(tempDate.getFullYear() + 3); //add 3 years
              endDate = tempDate.toISOString();
              if (tempDate > this.getMaxDate())
                endDate = this.getMaxDate().toISOString();
              //skipNumber resets to 0 for this next date batch
              return this.queryCalendar(calObj, calConfigs, existingEvent, 0, startDate, endDate);
            }
          }
        } //no error returned
      }) //end Graph.get()
      .catch(value => { //value is always undefined (need to save error details from with above .get func)
        //Just catch to prevent "Uncaught (in promise)" console error
        //Also needed so that Promise.all correctly resolves
      });
    }); //end Graph client
    else {
      return await new Promise<void>((reject) => { //resolve, reject
        //console.log("Flow 3 found, not calling for Outlook calendar")
        reject();
      });
    }
  } //end queryCalendar()
}
