import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITimelineCalendarProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  instanceId: string;
  //renderLegend: any;
  //setGroups: any;
  groups: any[];
  categories: any[];
  lists: any[];
  //renderEvents: any;
  getDefaultTooltip: any;
  buildDivStyles: any;
  context: WebPartContext;
  domElement: HTMLElement;
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
