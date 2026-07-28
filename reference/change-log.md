---
description: See changes per version
---

# Change log

## 2.8.9

* **New** [**cloud systems**](../developing-with-xlconnect/cloud-systems.md)**:**&#x20;
  * Nmbrs
  * AccountView

## 2.8.8

* **Partner model**, the ribbon tab now shows the partner name instead of "XLConnect" to allow partners to have a branded 'shop' for their content in their users' Excel.&#x20;
* **Cleaned up UI** buttons that don't work for a user because they do not have the required roles are hidden resulting in a cleaner UI for both normal users and developers.&#x20;
* **Re-order scripts** by drag & drop much requested option to re-order scripts
* [**xlc.requireScope**](the-xlc-object.md#scopes) script command to add ensure the scopes you need for your app are added without requiring user involvement&#x20;
* **Default table style** set to **Light 9** to prevent the less attractive default table style
* **New** [**cloud systems**](../developing-with-xlconnect/cloud-systems.md)**:**&#x20;
  * Payhawk
* **Metadata** now only stored within workbook. This an internal change that should not affect users, except this solves the edge cases where metadata got dropped when users add a new first tab in an Excel without XLConnect present. &#x20;
* **Dynamic Array connections** can now be saved to the cloud server.&#x20;

## 2.8.6

* **New cloud systems**:
  * QuickBooks Online
  * Exact Online
  * Visma Client Credentials
* **New** [**xlc.progress**](the-xlc-object.md#xlc.progress-int-value-1-string-message-null-int-max-0) method for simplified progress reporting

## 2.8.5

* Bugfix for xlc.put&#x20;

## 2.8.4



## 2.8.3



## 2.8.2



## 2.8.1



## 2.8.0

* **Autolink tables**: when a script has un unconnected table connection, i will look for a table with that name in the workbook and connect to that. This feature can be used to dynamically create tables in VBA and script and have the data flow to them.&#x20;
* **Prefabs**: insert prefab elements into yor workbook to kcikstart development. If you need to exapnd on their functionality you can edit the javascript. Please check the Insert Prefab regularly as new content for it will be released.
* **xlc.http method**: Allows for fine-grained http calls that also returns the reply headers.
* Bug fixes:
  * Sometimes the metadata could be deleted from workbooks, this has been fixed.&#x20;
  * JsonViewer could throw errors for certain drill down actions
* **SpillConnection**: Connect to dynamic arrays&#x20;

## 2.7.8

* **Named Connections**: when you need muliple connections to the same platform (because you have multiple tenants where the tenant gets coded into the token) you can now optionally specify a name for the connection. Say 'visma.net:NL' and 'visma.net:UK' these cookies will be stored separately so you can have multiple connecitons into your cloud system.&#x20;
* **Basic Authentication**: use "basic:name" as the authentication value for http.get etc, XLC will ask for your username and password and store them encrypted just like the cookies to other platforms.&#x20;
* **QuickView module function code**: double-click a function in a module in the Engine pane of the Javascript Studio to see the implementation. This allows you to quickly see how arguments are used.
* **Dynamic Start Menu**: right click a workbook in the Launch Pad and select 'Set as Start Menu' to have users open that workbook as their start menu.&#x20;
* **Desktop** and **Desktop.Standalone:** Run the Javascript Studio outside of Excel, or even standalone triggered by the Windows Task Scheduler to run ETL processes automatically.&#x20;

##

