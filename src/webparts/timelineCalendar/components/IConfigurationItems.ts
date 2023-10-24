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