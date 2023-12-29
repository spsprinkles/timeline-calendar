import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from '@microsoft/sp-http';

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
  calendars: any[];
  //renderEvents: any;
  getDefaultTooltip: any;
  buildDivStyles: any;
  context: WebPartContext;
  graphClient: Promise<MSGraphClientV3>;
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
