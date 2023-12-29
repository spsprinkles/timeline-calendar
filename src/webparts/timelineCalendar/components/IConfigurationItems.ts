import { IPersonaSharedProps } from '@fluentui/react/lib/Persona'; //Will "overwrite" IPersonaProps from this

//"Overwrite" the expected interface to add properties
export interface IPersonaProps extends IPersonaSharedProps {
    key: string
    mail: string
    personaType: "user" | "group"
}

export interface IMemberOfResult {
    "@odata.type": string
    id: string
    displayName?: string
    mail: string
    visibility: string
}

export interface ICalendarItem {
    uniqueId: string
    sortIdx?: number //auto field
    persona: IPersonaProps[]
    calendar: any
    category: string
    classField: string
    className: string
    group:string
    groupId: string
    groupField: string
    ignorePrivate: boolean
    visible: boolean
}

export interface ICategoryItem {
    uniqueId: string
    sortIdx?: number //auto field
    name: string
    borderColor: string
    bgColor: string
    textColor: string
    visible: boolean
    advancedStyles: string
}

export interface IGroupItem {
    uniqueId: string
    sortIdx?: number //auto field
    name: string
    visible: boolean
    html: string
    className: string
}

export interface IListItem {
    uniqueId: string
    sortIdx?: number //auto field
    siteUrl: string
    list: string
    listName: string //not filled by user
    isCalendar: boolean //not filled by user
    view: string
    viewName: string //not filled by user
    viewFilter: string //not filled by user
    titleField: string
    startDateField: string
    endDateField?: string
    category: string
    classField: string
    className: string
    group:string
    groupId: string
    groupField: string
    configs?: string //Advanced Configurations (which will encompass the below)
    visible?: boolean //advanced prop
    camlFilter?: string //advanced prop
    dateInUtc?: boolean //advanced prop
}