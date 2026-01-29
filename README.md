# Timeline Calendar

## ℹ️ Summary

![SPFx 1.17.2](https://img.shields.io/badge/SPFx-1.17.2-green.svg)
![Node.js v16](https://img.shields.io/badge/Node.js-v16-green.svg)
![SPO](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)

The Timeline Calendar web part can pull in data from multiple sources and render them together in a dynamic, filterable timeline view. This includes **SharePoint** lists (and "classic" calendars) as well as **Outlook** calendars (including **Microsoft 365 Group** calendars). Options are available to easily adjust how the timeline behaves, including configuring the tooltip, and there is support for making advanced configuration changes to refine the look and behavior of the timeline.

![Timeline Calendar web part](https://github.com/spsprinkles/timeline-calendar/assets/8918397/27d7632c-170e-443e-8b69-7d16fa6c3184)

## 📖 User Guide

Access the [**user/setup guide**](https://github.com/spsprinkles/timeline-calendar/wiki) for information on the specific options/properties available for customization. This includes details on both basic and advanced uses of the web part.

## 🔑 Graph API Permissions

As of version 0.6.0, the following [**delegated**](https://learn.microsoft.com/en-us/graph/permissions-overview#delegated-permissions) Graph API permissions/scopes are requested by the application. Failure to approve these permissions in the SPO Admin Center (or Azure Portal) will result in degraded functionality as specified below.

| Permission | API             | Reason        |
| ---------- | --------------- | ------------- |
| User.Read  | [/me/memberOf](https://learn.microsoft.com/en-us/graph/api/user-list-memberof) | Get list of M365 Groups for the _current_ user |
| User.Read.All | [/users](https://learn.microsoft.com/en-us/graph/api/user-list) | Search the directory for users & shared mailboxes (for the "people picker") |
| Group.Read.All | [/groups](https://learn.microsoft.com/en-us/graph/api/group-list) | Search the directory for M365 Groups (for the "people picker") and get group calendar events that the _current_ user has access to |
| Calendars.Read.Shared | [/users/${userId}/calendars](https://learn.microsoft.com/en-us/graph/api/user-list-calendars) | Query user's calendars that have been shared with the _current_ user |

## 📃 Content Security Policy (CSP)

Loading external scripts (such as from CDNs) within SharePoint Online is [subject to CSP enforcement](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/content-securty-policy-trusted-script-sources). The Timeline Calendar uses the Microsoft developed [Monaco Editor](https://github.com/microsoft/monaco-editor) for JSON & other configuration editing. It attempts to load the editor from the following CDN locations in the following order (except for USSec environment which is handled differently):

1. https://cdnjs.cloudflare.com/

2. https://cdn.jsdelivr.net/

SharePoint tenant administrators will need to follow the [Managing the Content Security Policy rules in SharePoint Online](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/content-securty-policy-trusted-script-sources#managing-the-content-security-policy-rules-in-sharepoint-online) section of the CSP article for guidance on adding one (or both) of the above locations (with the trailing `/` character) in the "Trusted script sources" area of the admin portal.

If the above locations are too "broad" for your tastes, you can scope them further to just the Monaco Editor path as shown below. Note that if you are adding the jsDelivr location, you must include the specific version as shown below. Again, it is important that you include the trailing `/` character:

1. https<nolink></nolink>://cdnjs.cloudflare.com/ajax/libs/monaco-editor/

2. https<nolink></nolink>://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/

> [!NOTE]
> The above are *not* addresses that you can open in a browser, but are rather the address of the base container/folder of the script files that get loaded.

If no location is added to the "Trusted script sources" area, the Monaco Editor will fail to load and a simple textarea editor (which lacks syntax highlighting) will be shown to users as a fallback (as of version 0.6.5).

## 📈 Version history

Refer to the [releases page](https://github.com/spsprinkles/timeline-calendar/releases) for specific details.

| Version | Date              | Comments        |
| ------- | ----------------- | --------------- |
| 0.6.5   | January 29, 2026  | Bug fixes for event duplication & box rendering and CSP fallback |
| 0.6.4   | January 17, 2025  | Graph & advanced editor updates + visual improvements |
| 0.6.3   | August 28, 2024   | Bug fixes for event querying |
| 0.6.2   | May 3, 2024       | Enabled advanced code editor for DoD365-Sec |
| 0.6.1   | March 18, 2024    | Several enhancements & bug fixes |
| 0.6.0   | January 21, 2024  | New features (including Outlook calendar support) & bug fixes |
| 0.5.3   | November 27, 2023 | Several bug fixes |
| 0.5.2   | October 20, 2023  | Initial release |

## ⚠️ Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
