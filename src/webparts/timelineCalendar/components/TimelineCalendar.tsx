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
import { ICalendarItem, ICategoryItem, IGroupItem, IListItem } from './IConfigurationItems';
import * as Handlebars from 'handlebars';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client'; //ResponseType
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
//import { DescriptionFieldLabel } from 'TimelineCalendarWebPartStrings';
//import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
//import { DefaultButton } from '@fluentui/react/lib/Button'; //PrimaryButton
//import { TeachingBubbleContentBase } from 'office-ui-fabric-react';

//declare const window: any; //temp TODO

class IdSvc {
  private static _id = 0;
  public static getNext(): number {
    this._id++;
    return this._id;
  }
  //private constructor() {}
}

// class Dispatcher {
//   private numOfCompletions = 0;
//   private warnings: string;
//   private errors: string;
// }

export default class TimelineCalendar extends React.Component<ITimelineCalendarProps, {}> {
  private _timeline: Timeline;
  private _dsItems: any;
  private _dsGroups: any;

  /**
   * Called when component is mounted (only on the *initial* loading of the web part)
   */
  public async componentDidMount(): Promise<void> {
    //const { data, calendars } = this.props;
    this.initialBuildTimeline();

    //Add a helper for when user mistypes a helper name so that an exception is not thrown
    Handlebars.registerHelper('helperMissing', function( /* dynamic arguments */) {
      if (arguments.length == 1) {
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
        if (strText == "1")
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
      if ((prevProps.groups == null || prevProps.groups.length == 0) && (this.props.groups && this.props.groups.length > 0)) { // && this.props.lists
        const groupId = (this.props.groups[0] as IGroupItem).uniqueId;
        const itemEvents = this._dsItems.get({ //get all events (except for "weekends")
          filter: function (item:any) {
            if (item.className != "weekend") {
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
          if (prevProps.holidayCategories != null && prevProps.holidayCategories != "" && item.className == self.ensureValidClassName(prevProps.holidayCategories)) {
            item.type = "range"; //assume it should be reverted to range (vs. point)
            item.group = groupId;
            return true;
          }
          else if (item.className == self.ensureValidClassName(self.props.holidayCategories)) {
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
            if (this.props.singleDayAsPoint && (item.end == null || item.start.toLocaleDateString() == item.end.toLocaleDateString())) //localDate == "11/27/2023"
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
        let tcWin = this._timeline.getWindow();
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
          <div id={"legend-" + instanceId} style={{display:"none"}}></div>
          <div id={"timeline-" + instanceId} />
        </div>
        <div id={"bottomGroupsBar-" + instanceId} className='bottomGroupsBar'></div>
        <div id={"dialog-" + instanceId}></div>
      </div>
    )
  }

  //Pass in the objects to merge as arguments (for a deep extend, set the first argument to true)
  //Cannot use Object.assign(options, userOptions) instead because it doesn't do deep property adding
  //TODO: Use this? https://developer.mozilla.org/en-US/docs/Web/API/structuredClone
  private extend(...args: any[]):any {
    const self = this;
    let extended = {} as any;
    let deep = false;
    let i = 0;
    const length = arguments.length;
  
    // Check if a deep merge
    if (Object.prototype.toString.call(arguments[0]) === '[object Boolean]') {
      deep = arguments[0];
      i++;
    }
  
    // Merge the object into the extended object
    const merge = function (obj:any) {
      for (var prop in obj) {
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
      var obj = arguments[i];
      merge(obj);
    }
  
    return extended;
  };

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

  private formatDateFromSOAP(d:string):Date {//"2014-08-28 23:59:00" or if UTC "2019-12-13T05:00:00Z"
    if (d == null)
			return null;
		
		//Check for UTC/Zulu time
		if (d.indexOf("Z") != -1)
			return new Date(d);
		else
			return new Date(d.replace(" ", "T")); //needed for IE
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
        if (this.props.tooltipEditor == null || this.props.tooltipEditor == "") {
          const handleTemplate = Handlebars.compile(this.props.getDefaultTooltip());
          return handleTemplate(item);
        }
        else {
          const handleTemplate = Handlebars.compile(this.props.tooltipEditor);
          return handleTemplate(item);
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
    let viewStart = new Date();
    viewStart.setDate(viewStart.getDate() - (this.props.initialStartDays || 7));
    return viewStart;
  }
  private getViewEndDate(): Date {
    //Set initial end view; default 3 months out
    let viewEnd = new Date();
    //viewEnd.setMonth(viewEnd.getMonth() + 3);
    viewEnd.setDate(viewEnd.getDate() + (this.props.initialEndDays || 90));
    return viewEnd;
  }

  private getMinDate(): Date {
    //Build dates for min/max data querying; default 2 months before today
    let minDate = new Date();
    minDate.setDate(minDate.getDate() - (this.props.minDays || 60));
    return minDate;
  }

  private getMaxDate(): Date {
    //Default max is 1 year from today; ensure the time is at the very end of the day (add a day but remove 1 second to get at the very end of the previous day)
    const now = new Date();
		let maxDate = new Date(now.getFullYear(), now.getMonth(), now.getDate()+1, 0, 0, -1);
    maxDate.setDate(maxDate.getDate() + (this.props.maxDays || 365));
    return maxDate;
  }
  
  private getGroupIdAtIndex(startingIndex:number): number {
    startingIndex = startingIndex || 0;
    let foundGroupId;

    const groups = this._dsGroups.get({
      order: "order",
      filter: function (item:any) {
        return (item.visible != false);
      }
    });
    for (var i=0; i<groups.length; i++) {
      if (i == startingIndex) {
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
      if (props.items.length > 0 && props.event.type == "tap") { //ignore the follow-on "press" event (still needed?)
				const oEvent = this._dsItems.get(props.items[0]);
        if (oEvent.encodedAbsUrl) {
					const itemUrl = oEvent.encodedAbsUrl.substring(0, oEvent.encodedAbsUrl.lastIndexOf("/")); //cut off the ending: "/ID#_.000"
          //OpenPopUpPage(itemUrl + "/DispForm.aspx?ID=" + oEvent.spId, function(result) {
					// 	if (result == 0 || oEvent.spId.toString().contains("T")) //Edits to series/recurring events with IDs like "16.0.2020-06-08T16:00:00Z" actually have a different ID generated by the SP form action
					// 		return;
					// 	//Handle item updates or deletions
          //}
          //Just open in new tab for now
          window.open(itemUrl + "/DispForm.aspx?ID=" + oEvent.spId, "_blank");

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
      }
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
      let labelElem = GetVisLabelElement(e.target as HTMLElement);
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
      let labelElem = GetVisLabelElement(e.target as HTMLElement);
      if (labelElem) {
        e.preventDefault(); //stop the normal menu from appearing

        //Is this the only group currently being shown?
        const shownGroups = this._dsGroups.get({
          filter: function (group:any) {
            return (group.visible != false);
          }
        });

        if (shownGroups.length == 1) {
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
            return (group.id != foundGroupId && group.visible != false);
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
      let childElem = e.target as HTMLDivElement;
      if (childElem.classList.contains('vis-item')) {
        processBottomBarItem(childElem);
        
        //Are there any group boxes still inside the bottomBar?
        if (bottomGroupsBar.querySelectorAll(".vis-item").length == 0)
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
    if (true) {//this.props.shadeWeekends) {
      let weekendStart = null as Date;
      if (this.options.min.getDay() == 0) { //0 = Sunday
        //Need to set weekend one day before
        weekendStart = new Date(this.options.min.getFullYear(), this.options.min.getMonth(), this.options.min.getDate() - 1);
      }
      else if (this.options.min.getDay() == 6) { //6 = Saturday
        weekendStart = new Date(this.options.min.valueOf());
      }
      else {
        //Add days to get to Saturday
        const daysToAdd = 6 - this.options.min.getDay();
        weekendStart = new Date(this.options.min.getFullYear(), this.options.min.getMonth(), this.options.min.getDate() + daysToAdd);
      }
      //Add two days but remove 1 second to get at the very end of the previous day
      let weekendEnd = new Date(weekendStart.getFullYear(), weekendStart.getMonth(), weekendStart.getDate()+2, 0, 0, -1);
      
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
    if (styleElem == null) {
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
        if (this.props.hideItemBoxBorder && divStyles.indexOf("background-color") == -1) //set bg to border color, mostly only applicable for vertical Holidays
          divStyles += "background-color:" + categoryItem.borderColor + ";";
        styleHtml += '.vis-item.' + this.ensureValidClassName(categoryItem.name) + ' {' + divStyles + '}\r\n';

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

    if (this.props.cssOverrides)
      styleHtml += "/* CSS class overrides */\r\n" + this.props.cssOverrides; //built-in protections for returning/encoding <script> as \x3Cscript>

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
          newElem.className = "legendBox vis-item vis-range " + this.ensureValidClassName(value.name);
          newElem.dataset.className = this.ensureValidClassName(value.name);
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
              return (item.className && item.className.split && item.className.split(" ").indexOf(className) != -1);
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
      let childElem = e.target as HTMLDivElement;
      if (childElem.classList.contains('legendBox')) {
        e.preventDefault(); //stop the normal menu from appearing
        
        //Is this the only one being shown?
				//var $legendBoxes = $("#legend > .legendBox:not(.gray):not(.print)");
        const activeLegendBoxes = legendBar.querySelectorAll(".legendBox:not(.gray)");
        if (activeLegendBoxes.length == 1 && childElem.classList.contains("gray") == false) {
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
          if (childElem != elem)
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
      if (groupsFormatted.length == 0)
        this._timeline.setGroups(null); //needed to get events to render if there are no groups
      else {
        this._dsGroups.add(groupsFormatted);
        this._timeline.setGroups(this._dsGroups);
      }
    }
  }

  private renderEvents(): void {
    //Function: showLegend (must be delared before/above where it's called)
    const showLegend = () => {
      //Hide the loader image
      document.getElementById("loading-" + this.props.instanceId).style.display = "none";
      //Show the legend
      document.getElementById("legend-" + this.props.instanceId).style.display = "block";
    }

    //Remove any existing events to prevent duplicate event adding (while in edit mode)
    const itemEvents = this._dsItems.get({
      filter: function (item:any) {
        return (item.className != "weekend");
      }
    });
    this._dsItems.remove(itemEvents);

    //Get SharePoint list/calendar data
    let spPromise = null as Promise<void | any[]>;
    if (this.props.lists) {
      //Get the view CAML
      spPromise = this.getViewsCAML().then(() =>{
        //Now get the events
        this.getSharePointEvents().then(response => {
          //console.log("all data returned, response is undefined because no data is actually returned");
          //showLegend();
        });
      });
    }

    //Get Outlook calendar events
    let outlookPromise = null as Promise<void | any[]>;
    if (this.props.calendars) {
      outlookPromise = this.getOutlookEvents()
    }

    //When both are finished
    Promise.all([spPromise, outlookPromise]).then(response => {
      //console.log("all data returned, response is undefined because no data is actually returned");
      showLegend();
    });
  }

  //private async, return await
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
            data.value.forEach((view:any) => {
              if (list.view.toLowerCase() == view.Title.toLowerCase() || list.view.toLowerCase().indexOf(view.Id) != -1) {
								//TC.log("Got ViewQuery for '" + view.Title + "' view");
								//view.ViewQuery "<Where><And><DateRangesOverlap><FieldRef Name="EventDate" /><FieldRef Name="EndDate" /><FieldRef Name="RecurrenceID" /><Value Type="DateTime"><Month /></Value></DateRangesOverlap><Eq><FieldRef Name="Category" /><Value Type="Text">Birthday</Value></Eq></And></Where>"
								//cal.viewFilter = view.ViewQuery.replace("<Month />", "<Year />");
								
								//Extract just the CAML filter portion
								const droIndex = view.ViewQuery.indexOf('</DateRangesOverlap>');
								const endIndex = view.ViewQuery.lastIndexOf('</And></Where>');
								const whereStartIndex = view.ViewQuery.indexOf('<Where>');
								const whereEndIndex = view.ViewQuery.lastIndexOf('</Where>');
								//Look for Calendar & Standard w/ Recurrence views (they have <DateRangesOverlap>)) //or add view.ViewType: "CALENDAR"
								if (droIndex > -1 && endIndex > -1)
									list.viewFilter = view.ViewQuery.substring(droIndex+20, endIndex);
								else if (whereStartIndex > -1)
									list.viewFilter = view.ViewQuery.substring(whereStartIndex+7, whereEndIndex);
								
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

  private getFieldKeys(input?:string):string[] {
    let fieldKeys = [] as string[];
    let source = (input || this.props.tooltipEditor);
    if (source) {
      //Extract the {{property}} references
      fieldKeys = source.match(/{{(.*?)}}/g);
      if (fieldKeys) {
        //Remove the vis.js "default" fields
        fieldKeys = fieldKeys.filter(i => {
          if (i != "{{content}}" && i != "{{start}}" && i != "{{end}}")
              return i;
        });
        //Extract just the field/"property" text from inside the {{ }} or {{{ }}}
        for (let i=0; i < fieldKeys.length; i++) {
            const matchResults = fieldKeys[i].match(/\w+/g);
            if (matchResults.length == 1)
              fieldKeys[i] = matchResults[0];
            else //handle cases like "{{{limit Description}}}" where the actual Description property is at the end of the match
              fieldKeys[i] = matchResults[1];
        }
      }
    }
    return fieldKeys;
  }

  private getSharePointEvents(): Promise<any[]> {
    return Promise.all(this.props.lists.map((list:IListItem) => {
      //Check for advanced configs
      if (list.configs && list.configs.trim() != "") {
        try {
          const configs = JSON.parse(list.configs);
          if (configs.camlFilter && configs.camlFilter.trim() != "")
            list.camlFilter = configs.camlFilter;
          
          //Set visible prop
          list.visible = (configs.visible == false ? false : true);

          //Set UTC prop
          if (configs.dateInUtc != null)
            list.dateInUtc = configs.dateInUtc;
          
          //Option to dynamically assign/build
          //const userOptions = JSON.parse(this.props.visJsonProperties); //just in case
          //options = this.extend(true, options, userOptions); //userOptions override set "defaults" above
        }
        catch (e) {}
      }
      

      if (list.visible == false)
        return; //skip this one

      return this.queryList(list);
    }))
  }

  private calendarEventsSoapEnvelope(list:IListItem, nextPageDetail?:string):string {
    const fieldKeys = this.getFieldKeys();
    const includeRecurrence = (list.isCalendar != false);
    
    //Set classField and className props
    //Split on the : char to determine if a field or category was selected (Field:fieldInternalName or Static:category.uniqueId)
    if (list.category) {
      const catValues = list.category.split(":");
      if (catValues[0] == "Field") {
        list.classField = catValues[1];
        if (list.className)
          list.className = null;
      }
      else { //[0] assumed to be "Static"
        const categoryId = catValues[1]; //Will be the uniqueId, need to get the display name next
        if (this.props.categories) {
          this.props.categories.every((category:ICategoryItem) => {
            if (category.uniqueId == categoryId) {
              list.className = category.name; //store the display name instead
              if (list.classField)
                list.classField = null;
              return false; //exit
            }
            else return true; //keep looping
          });
        }
      }
    }

    //Set groupField and groupId props
    //Split on the : char to determine if a field or category was selected
    if (list.group) {
      const catValues = list.group.split(":");
      if (catValues[0] == "Field") {
        list.groupField = catValues[1];
        if (list.groupId)
          list.groupId = null;
      }
      else { //[0] assumed to be "Static"
        list.groupId = catValues[1]; //Will be the uniqueId
        if (list.groupField)
          list.groupField = null;
      }
    }

    let returnVal = "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/'><soapenv:Body><GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" + 
			"<listName>" + list.list + "</listName>" + 
			"<viewFields><ViewFields>" + 
				"<FieldRef Name='Title' />" + 
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
        (list.classField ? "<FieldRef Name='" + list.classField + "' />" : "") +
        (list.groupField ? "<FieldRef Name='" + list.groupField + "' />" : "");

    //Add fields used in the tooltip
    fieldKeys.forEach(field => {
      returnVal += "<FieldRef Name='" + field + "' />";
    });
    
    //Add non-calendar fields
    if (list.startDateField) //list.isCalendar == false &&
      returnVal += "<FieldRef Name='" + list.startDateField + "' />";
    if (list.endDateField)
      returnVal += "<FieldRef Name='" + list.endDateField + "' />";
    
    returnVal += "</ViewFields></viewFields>" + 
    //Set rowLimit, without this the default is only 30 items
    "<rowLimit>1000</rowLimit>"; //Setting as 0 initially *seemed* to get all events but SP would stop at a seemingly random point
    //Even with setting this to 5000 SP only returns 999 events and requires you to use pagination for next batch of events
    
    let query = "<query><Query><Where>"; //<Where> is the start of the ViewQuery property from calendar views

    //if (cal.filter != "") //this was needed to prevent the general filter from applying when no calendar filter is wanted
    //	query += (cal.filter ? "<And>" + cal.filter : (TC.settings.filter ? "<And>" + TC.settings.filter : ""));

    //Determine applicable filter
    let filterToUse = null;
    if (list.viewFilter)
      filterToUse = list.viewFilter;
    else if (list.camlFilter && list.camlFilter.trim() != "")
      filterToUse = list.camlFilter;

    if (filterToUse)
      query += "<And>" + filterToUse;
    
    //Handle custom lists/non-calendars
    if (list.isCalendar == false) {
      //Was an end date field provided?
      if (list.endDateField) //use both dates from the non-calendar
      query += "<Or><Geq><FieldRef Name='" + list.startDateField + "'/><Value Type='DateTime'>" + this.getMinDate().toISOString() + "</Value></Geq>" +
        "<Geq><FieldRef Name='" + list.endDateField + "'/><Value Type='DateTime'>" + this.getMinDate().toISOString() + "</Value></Geq>" +
      "</Or>";
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
      //"<CalendarDate>2015-08-22T12:00:00.000Z</CalendarDate>" + //today is the default if no element provided
      (nextPageDetail ? "<Paging ListItemCollectionPositionNext='" + nextPageDetail + "' />" : "" ) +
      (includeRecurrence ? "<RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion>" : "") + 
      (includeRecurrence ? "<ExpandRecurrence>TRUE</ExpandRecurrence>" : "") + 
      (includeRecurrence ? "<RecurrenceOrderBy>TRUE</RecurrenceOrderBy>" : "") +
      "<ViewAttributes Scope='RecursiveAll' />" +
      //"<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" +
      "<ViewFieldsOnly>TRUE</ViewFieldsOnly>" +
      //(this.props.getDatesAsUtc ? "<DateInUtc>TRUE</DateInUtc>" : "") + //True returns dates as "2023-10-10T06:00:00Z" versus "2023-10-10 08:00:00"
      (list.dateInUtc == false ? "" : "<DateInUtc>TRUE</DateInUtc>") +
    "</QueryOptions></queryOptions></GetListItems></soapenv:Body></soapenv:Envelope>";
    
    return returnVal;
  }

  private handleMultiValues(fieldValue:string): string[] {
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
    if (fieldSplit.length == 2 && isNaN(Number(fieldSplit[0]))) {
      //Remove the first value from the array (leaving a single entry array)
      fieldSplit.shift();
    }
    
    //Check for multiple-value fields 
    //Remove potential blank entries in split (from multi-choice fields)
    fieldSplit = fieldSplit.filter(i => {return i});

    //A single value was found (either a "regular" value or from a Calculated field or a multi-choice field)
    if (fieldSplit.length == 1) {
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
          return index % 2 != 0; //Index is odd
        });
      }
    }

    return fieldSplit;
  }

  private async queryList(list:IListItem, nextPageDetail?:string): Promise<void> {
    return await this.props.context.spHttpClient.post(list.siteUrl + "/_vti_bin/lists.asmx", SPHttpClient.configurations.v1,
    {
      headers: [
        ["Accept", "application/xml, text/xml, */*; q=0.01"],
        ["Content-Type", 'text/xml; charset="UTF-8"']
      ],
      body: this.calendarEventsSoapEnvelope(list, nextPageDetail)
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
      //let numOfValidItems = 0;
      let lastStartDate:Date = null;
      const fieldKeys = this.getFieldKeys();
      
      //Need to determine which date fields to use (calendar vs non-calendars)
      const startFieldName = (list.isCalendar == false ? "ows_" + list.startDateField : "ows_EventDate"); //must have a start date
      const endFieldName = (list.isCalendar == false ? (list.endDateField ? "ows_" + list.endDateField : null) : "ows_EndDate");

      let pagingDetails:string = null;
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(strXml, "application/xml");
      xmlDoc.querySelectorAll("*").forEach(elem => {
        if (elem.nodeName == "rs:data")
          //Store this for use after the forEach
          pagingDetails = elem.getAttribute("ListItemCollectionPositionNext"); //ItemCount is another
        
        //Loop over the event/data results
        else if(elem.nodeName == "z:row") { //actual data is here
          //Build date variables
          const allDayEvent = (elem.getAttribute("ows_fAllDayEvent") == "1" ? true : false);
          let strStartDateValue = elem.getAttribute(startFieldName);
          if (allDayEvent)
            strStartDateValue = strStartDateValue.split("Z")[0]; //Drop zulu to make it handle date as local
          const eventStartDate = this.formatDateFromSOAP(strStartDateValue);

          let eventEndDate;
          if (endFieldName) {
            let strEndDateValue = elem.getAttribute(endFieldName);
            if (allDayEvent)
              strEndDateValue = strEndDateValue.split("Z")[0]; //Drop zulu to make it handle date as local
            eventEndDate = this.formatDateFromSOAP(strEndDateValue);
            //Check for non-calendar list dates in which no time is provided...
            if (eventEndDate.getHours() == 0 && eventEndDate.getMinutes() == 0) {
              //...change the time to the end of day
              eventEndDate.setDate(eventEndDate.getDate() + 1);
              eventEndDate.setSeconds(eventEndDate.getSeconds() - 1);
              //strEndDate = eventEndDate.format("yyyy-MM-dd HH:mm:ss");
              //These are not the UTC versions...
              //strEndDate = eventEndDate.getFullYear().toString() + "-" + (eventEndDate.getMonth()+1).toString() +
            }
          }
          lastStartDate = eventStartDate; //saved for later
          
          //CAML returns *recurring* events for a whole year prior to and after today regardless of any EventDate/EndDate filters
          //So only include events that fall within the requested date range
          if ((eventEndDate == null || eventEndDate >= this.getMinDate()) && eventStartDate <= this.getMaxDate()) {
            //numOfValidItems++;
            
            //Get the "title" value
            let strTitle = "[No Title]"; //default value, changed next
            if (list.titleField) {
              strTitle = (elem.getAttribute("ows_" + list.titleField) || strTitle);
              //Check for and handle calculated & lookup column formatting
              const splits = strTitle.split(";#"); //string;#Custom title
              strTitle = (splits[1] || splits[0]); //0 index is for other/"regular" fields
            }
            else
              strTitle = (elem.getAttribute("ows_Title") || strTitle);
              
            //Build the event obj
            let oEvent = {
              id: IdSvc.getNext(),
              spId: elem.getAttribute("ows_ID"),
              encodedAbsUrl: elem.getAttribute("ows_EncodedAbsUrl"),
              content: strTitle,
              //title: elem.getAttribute("ows_Title"), //Tooltip
              start: eventStartDate,
              //end: elem.getAttribute("ows_EndDate"),
              type: "range", //Changed later as needed
              //className: //assigned next
              //group: list.groupId //assigned next
            } as any;
            
            //Add end date if applicable
            if (eventEndDate) {
              oEvent.end = eventEndDate;
              //Force single day events as point?
              if (this.props.singleDayAsPoint && (elem.getAttribute(startFieldName).substring(0, 10) == elem.getAttribute(endFieldName).substring(0, 10)))
                oEvent.type = "point";
            }
            else //no end date, so make it a point
              oEvent.type = "point";            

            //Add class/category
            if (list.className)
              oEvent.className = this.ensureValidClassName(list.className);
            //Apply class by category field
            else if (list.classField)
              oEvent.className = this.ensureValidClassName(elem.getAttribute("ows_" + list.classField));

            /* Handled above instead, and address where end date is after start but still has no time
            //Special checks for range events (mostly for non-calendar lists)
            if (oEvent.type == "range") {
              //"range" events needs an end property, but also make sure "same day" events show ending at end of the day
              if (oEvent.end == null || oEvent.start.getTime() == oEvent.end.getTime()) {
                var end = new Date(oEvent.start.getTime());
                end.setDate(end.getDate() + 1);
                end.setSeconds(end.getSeconds() - 1);
                oEvent.end = end;
              }
            }*/

            // //Set holiday render format
            // if (oEvent.className == TC.settings.holidayClass && TC.settings.holidayType != "point") { //override render type to range
            //   oEvent.type = "range";
            //   //Vertical holiday?
            //   if (TC.settings.holidayType == "verticalBar") {
            //     /*oEvent.className += " verticalBar";
            //     //Force event to top group
            //     oEvent.group = topGroupIds.firstGroupId;*/
            //     oEvent.type = "background";
            //   }
            // }
            if (oEvent.className && oEvent.className == this.ensureValidClassName(this.props.holidayCategories)) {
              oEvent.type = "background"; //change to background
              oEvent.group = null; //apply to entire timeline
            }
            
            // //Call custom function if provided
            // if (TC.settings.beforeEventAdded)
            //   oEvent = TC.settings.beforeEventAdded(oEvent, $(this), cal);

            //Add data to the event object (for later tooltip template processing)
            fieldKeys.forEach(field => {
              //Skip these fields to prevent their above defined value from being overwritten
              if (field == "id" || field == "content" || field == "start" || field == "end" || field == "type" || field == "className")
                return;

              let fieldValue = elem.getAttribute("ows_" + field);
              if (fieldValue) {
                oEvent[field] = (fieldValue || ""); //save initial value
                //Look for special fields
                if (field == "Description") {
                  //Remove blanks
                  if (fieldValue == "<div></div>" || fieldValue == "<div></div><p>​</p>") //last one has a *hidden* character
                    fieldValue = "";
                  //If HTML text (versus plain text), starts with < character
                  if (fieldValue.indexOf("<") == 0) {
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
                      if (elem.style.backgroundColor == "#ffffff" || elem.style.backgroundColor == "rgb(255, 255, 255)" ||
                            elem.style.backgroundColor == "#dfdfdf")
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
                  oEvent[field] = fieldValue;
                }
                else {
                  //Handle Number/Currency field values like 1.00000000000000 & 10.2000000000000
                  //@ts-ignore endsWith is valid
                  if (isNaN(Number(fieldValue)) == false && fieldValue.endsWith("0000")) {
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
                  const fieldSplit = this.handleMultiValues(fieldValue);
                  oEvent[field] = fieldSplit.join(", ");
                }
              }
            });

            //Add group (row/swimlane)
            let multipleValuesFound = false;
            if (list.groupId)
              oEvent.group = list.groupId;
            else if (list.groupField && this.props.groups) {
              //Find the associated group to assign the item to
              let groupFieldValue = elem.getAttribute("ows_" + list.groupField);
              if (groupFieldValue) {
                const groupSplit = this.handleMultiValues(groupFieldValue);

                //Look for "regular" values without ;# (they result in ["Single value"] array)
                if (groupSplit.length == 1) {
                  groupFieldValue = groupSplit[0];
                }
                else {
                  //More than one, *real* value is in the field
                  multipleValuesFound = true;

                  //Create a duplicate event for each selected group value
                  groupSplit.forEach(groupName => {
                    const eventClone = structuredClone(oEvent); //error TS2304: Cannot find name 'structuredClone'
                    //Above duplicates the event object
                    eventClone.id = IdSvc.getNext(); //Set a new ID

                    //Find the associated group from it's name
                    this.props.groups.every((group:IGroupItem) => {
                      if (group.name == groupName) {
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

            //Add event/item to the DataSet
            if (multipleValuesFound == false)
              this._dsItems.add(oEvent);
          }
        }
      }); //end SOAP response forEach

      //Check if more data should be queried
      if (pagingDetails) {
        let validStartTime = true; //initial
        //Non-calendars seem to be returning these formats:
        // > Paged=Next&p_StartTimeUTC=00000101T588376552265528Z - this one is too long?
        // > Paged=Next&p_StartTimeUTC=18991230T000001Z - this one is invalid
        //Which even when using that in the next call it returns the same items. So prevent these calls from occuring by checking the "T" position
        const searchParams = new URLSearchParams(pagingDetails);
        const strTemp = searchParams.get("p_StartTimeUTC");
        //@ts-ignore (for startsWith)
        if (strTemp.indexOf("T") != 8 || strTemp.startsWith("1899")) {
          validStartTime = false;
          console.log(list.listName + " returned an invalid ListItemCollectionPositionNext: " + pagingDetails + " (ignoring)");
        }

        if (validStartTime && lastStartDate < this.getMaxDate()) {						
          //Need to keep "&" character encoded for next SOAP call
          pagingDetails = pagingDetails.replace("&p_", "&amp;p_");
          //Query for more events
          return this.queryList(list, pagingDetails);
        }
      }
    })
    .catch((error: any) => {
      console.error(error);
    });
  }

  private getOutlookEvents(): Promise<any[]> {
    return Promise.all(this.props.calendars.map((calendar:ICalendarItem) => {
      /*
      //Check for advanced configs
      if (list.configs && list.configs.trim() != "") {
        try {
          const configs = JSON.parse(list.configs);
          if (configs.camlFilter && configs.camlFilter.trim() != "")
            list.camlFilter = configs.camlFilter;
          
          //Set visible prop
          list.visible = (configs.visible == false ? false : true);

          //Set UTC prop
          if (configs.dateInUtc != null)
            list.dateInUtc = configs.dateInUtc;
          
          //Option to dynamically assign/build
          //const userOptions = JSON.parse(this.props.visJsonProperties); //just in case
          //options = this.extend(true, options, userOptions); //userOptions override set "defaults" above
        }
        catch (e) {}
      }
      */

      if (calendar.visible == false || calendar.persona == null || calendar.persona.length == 0)
        return; //skip this one

      return this.queryCalendar(calendar, 0);
    }))
  }
  
  //TODO: Compare with SP list; Removed private *async* and return *await*

  private queryCalendar(calObj:ICalendarItem, skipNumber:number): Promise<void> {
    //Set classField and className props
    //Split on the : char to determine if a field or category was selected (Field:owaField or Static:category.uniqueId)
    if (calObj.category) {
      const catValues = calObj.category.split(":");
      if (catValues[0] == "Field") {
        calObj.classField = catValues[1];
        if (calObj.className)
          calObj.className = null;
      }
      else { //[0] assumed to be "Static"
        const categoryId = catValues[1]; //Will be the uniqueId, need to get the display name next
        if (this.props.categories) {
          this.props.categories.every((category:ICategoryItem) => {
            if (category.uniqueId == categoryId) {
              calObj.className = category.name; //store the display name instead
              if (calObj.classField)
                calObj.classField = null;
              return false; //exit
            }
            else return true; //keep looping
          });
        }
      }
    }

    //Set groupField and groupId props
    //Split on the : char to determine if a field or category was selected
    if (calObj.group) {
      const catValues = calObj.group.split(":");
      if (catValues[0] == "Field") {
        calObj.groupField = catValues[1];
        if (calObj.groupId)
          calObj.groupId = null;
      }
      else { //[0] assumed to be "Static"
        calObj.groupId = catValues[1]; //Will be the uniqueId
        if (calObj.groupField)
          calObj.groupField = null;
      }
    }

    return this.props.graphClient.then((client:MSGraphClientV3): void => {
      const fieldKeys = this.getFieldKeys();
      //singleValueExtendedProperties doesn't need to be selected (the $expand seems to include it already)
      let selectFields = "id,createdDateTime,lastModifiedDateTime,categories,subject,isAllDay,webLink,body,start,end,location,organizer";
      //Add fields used in the tooltip
      fieldKeys.forEach(field => {
        //Only select Event properties can be allowed or error received:
        //HTTP 400: Could not find a property named 'Category' on type 'Microsoft.OutlookServices.Event'.
        switch(field) {
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
            selectFields += "," + field;
        }
      });

      //Build API URL based on user or group calendar
      let apiURL = "";
      if (calObj.persona[0].personaType == "user")
        apiURL = "/users/" + calObj.persona[0].mail + "/calendars/" + calObj.calendar + "/calendarView";
      else //assumed to be a group
        apiURL = "/groups/" + calObj.persona[0].key + "/calendarView";

      //Get calender view events
      client.api(apiURL).query(`startDateTime=${this.getMinDate().toISOString()}&endDateTime=${this.getMaxDate().toISOString()}`)
      .select(selectFields).top(500).skip(skipNumber)
      //TODO: .filter("sensitivity ne 'private'")
      //Make the charm/icon value be included
      .expand("singleValueExtendedProperties($filter=id eq 'Integer {11000E07-B51B-40D6-AF21-CAA85EDAB1D0} Id 0x0027')")
      //Headers -> Prefer: outlook.timezone (string) -> If not specified dates are returned in UTC
      .get((error:GraphError, response:any, rawResponse?:any) => {
        if (error) {
          //resolve(error.message);
        }
        else {
          //Extend Graph.Event interface with implicitly defined index signature (to support eventObj["keyFieldName"] retrieval)
          const events:(MicrosoftGraph.Event & {[key: string]:any})[] = response.value;
          events.forEach(calEvent => {
            //Get the "title" value
            let strTitle = "[No Title]"; //default value, changed next
            if (calEvent.subject != null && calEvent.subject.trim() != "") {
              strTitle = calEvent.subject;
            }
            
            //Process "non-dates" ("0001-01-01T00:00:00Z" is returned for private events)
            if (calEvent.createdDateTime == "0001-01-01T00:00:00Z") { //had: calEvent.sensitivity == "private" && 
              calEvent.createdDateTime = null;
              calEvent.lastModifiedDateTime = null;
            }

            //Build the event obj
            let oEvent = {
              id: IdSvc.getNext(),
              spId: calEvent.id, //TODO: Rename to sourceId?
              encodedAbsUrl: calEvent.webLink,
              content: strTitle,
              //title: elem.getAttribute("ows_Title"), //Tooltip
              start: new Date(calEvent.start.dateTime + (calEvent.isAllDay ? "" : "Z")), //treat as local time for all day events
              end: new Date(calEvent.end.dateTime + (calEvent.isAllDay ? "" : "Z")),
              type: "range", //Changed later as needed
                  //  location: calEvent.location.displayName,
              //className: //assigned next
              //group: list.groupId //assigned next
            } as any;
            
            //Force single day events as point?
            if (this.props.singleDayAsPoint && (calEvent.start.dateTime.substring(0, 10) == calEvent.end.dateTime.substring(0, 10)))
              oEvent.type = "point";

            //Add class/category
            if (calObj.className)
              oEvent.className = this.ensureValidClassName(calObj.className);
            //Apply class by category field
            else if (calObj.classField)
              oEvent.className = this.ensureValidClassName(calEvent[calObj.classField]);

            /* Handled above instead, and address where end date is after start but still has no time

            //Build date variables
            const allDayEvent = (elem.getAttribute("ows_fAllDayEvent") == "1" ? true : false);
            let strStartDateValue = elem.getAttribute(startFieldName);
            if (allDayEvent)
              strStartDateValue = strStartDateValue.split("Z")[0]; //Drop zulu to make it handle date as local
            const eventStartDate = this.formatDateFromSOAP(strStartDateValue);

            let eventEndDate;
            if (endFieldName) {
              let strEndDateValue = elem.getAttribute(endFieldName);
              if (allDayEvent)
                strEndDateValue = strEndDateValue.split("Z")[0]; //Drop zulu to make it handle date as local
              eventEndDate = this.formatDateFromSOAP(strEndDateValue);
              //Check for non-calendar list dates in which no time is provided...
              if (eventEndDate.getHours() == 0 && eventEndDate.getMinutes() == 0) {
                //...change the time to the end of day
                eventEndDate.setDate(eventEndDate.getDate() + 1);
                eventEndDate.setSeconds(eventEndDate.getSeconds() - 1);
                //strEndDate = eventEndDate.format("yyyy-MM-dd HH:mm:ss");
                //These are not the UTC versions...
                //strEndDate = eventEndDate.getFullYear().toString() + "-" + (eventEndDate.getMonth()+1).toString() +
              }
            }


            //Special checks for range events (mostly for non-calendar lists)
            */

            if (oEvent.className && oEvent.className == this.ensureValidClassName(this.props.holidayCategories)) {
              oEvent.type = "background"; //change to background
              oEvent.group = null; //apply to entire timeline
            }
            
            // //Call custom function if provided
            // if (TC.settings.beforeEventAdded)
            //   oEvent = TC.settings.beforeEventAdded(oEvent, $(this), cal);

            //Add data to the event object (for later tooltip template processing)
            /*
            Location (location.displayName)
            Category (categories[])
            Description (body.content assumed body.contentType == "html")
              "content": "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body>
              <div class=\"cal_1662 (number changes)\"></div>
              <div class=\"cal_1662\">
                <div>
                  <div class=\"x_elementToProof elementToProof\" style=\"font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)\">
                    Just entering text produces this. Now edited by Mike.</div></div></div></body></html>"
            Author ( "organizer": { "emailAddress": { "name": "John Doe",)
            Editor (? perhaps an extended MAPI property)
            Modified (lastModifiedDateTime) //and Created (createdDateTime) as "ISOZ" string
            */
            fieldKeys.forEach(field => {
              //Skip these fields to prevent their above defined value from being overwritten
              if (field == "id" || field == "content" || field == "start" || field == "end" || field == "type" || field == "className")
                return;

              let fieldValue = calEvent[field]; //TODO: support object values like organizer.emailAddress.address
              let wasMapped = true;
              //Map certain field names
              switch (field) {
                case "Location":
                  fieldValue = calEvent.location.displayName;
                  oEvent["Location"] = fieldValue;
                  break;

                case "Category":
                  fieldValue = calEvent.categories.join(", ");
                  oEvent["Category"] = fieldValue;
                  break;

                case "Description":
                  fieldValue = (calEvent.body && calEvent.body.content || "");
                  oEvent["Description"] = fieldValue;
                  break;

                case "Author":
                  fieldValue = (calEvent.organizer && calEvent.organizer.emailAddress.name || "");
                  oEvent["Author"] = fieldValue;
                  break;

                // case "Editor":
                //   fieldValue = calEvent.??; //perhaps an extended MAPI property
                //   oEvent["Editor"] = fieldValue;
                //   break;

                case "Created":
                  fieldValue = calEvent.createdDateTime;
                  oEvent["Created"] = fieldValue;
                  break;
                
                case "Modified":
                  fieldValue = calEvent.lastModifiedDateTime;
                  oEvent["Modified"] = fieldValue;
                  break;

                case "charmIcon":
                  fieldValue = (calEvent.singleValueExtendedProperties && calEvent.singleValueExtendedProperties[0] &&
                    calEvent.singleValueExtendedProperties[0].value || "");
                  if (fieldValue == "None")
                    fieldValue = "";
                  oEvent["charmIcon"] = fieldValue;
                  break;

                default:
                  wasMapped = false;
              }

              if (fieldValue && wasMapped == false) {
                oEvent[field] = (fieldValue || ""); //save initial value
              }
            });

            //Add group (row/swimlane)
            let multipleValuesFound = false;
            if (calObj.groupId)
              oEvent.group = calObj.groupId;
            else if (calObj.groupField && this.props.groups) {
              //Get value of group field
              let groupFieldValue = null as string | string[];
              switch (calObj.groupField) {
                case "categories":
                  groupFieldValue = calEvent.categories;
                  break;

                case "showAs":
                  //MicrosoftGraph.FreeBusyStatus = "unknown" | "free" | "tentative" | "busy" | "oof" | "workingElsewhere";
                  groupFieldValue = calEvent.showAs.toString();
                  break;

                case "charm":
                  groupFieldValue = (calEvent.singleValueExtendedProperties && calEvent.singleValueExtendedProperties[0] &&
                    calEvent.singleValueExtendedProperties[0].value || "");
                  if (groupFieldValue == "None")
                  groupFieldValue = "";
                  break;

                default:
                  groupFieldValue = calEvent[calObj.groupField];
              }

              //Find the associated group to assign the item to
              if (groupFieldValue) {
                if (Array.isArray(groupFieldValue)) {
                  multipleValuesFound = true;

                  //Create a duplicate event for each selected group value
                  groupFieldValue.forEach(groupName => {
                    const eventClone = structuredClone(oEvent); //error TS2304: Cannot find name 'structuredClone'
                    //Above duplicates the event object
                    eventClone.id = IdSvc.getNext(); //Set a new ID

                    //Find the associated group from it's name
                    this.props.groups.every((group:IGroupItem) => {
                      if (group.name == groupName) {
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

            //Add event/item to the DataSet
            if (multipleValuesFound == false)
              this._dsItems.add(oEvent);
          }); //foreach event

          //Check if more data should be queried
          let nextLink = response["@odata.nextLink"] as string;
          if (nextLink) {
            //Query for more events (get the next page)
            const eqIndex = nextLink.lastIndexOf("=");
            const skipNumber = Number(nextLink.substring(eqIndex + 1));
            return this.queryCalendar(calObj, skipNumber);
          }
        } //no error returned
      }); //end Graph.get()
    }); //end Graph client
  } //end queryCalendar()
}
