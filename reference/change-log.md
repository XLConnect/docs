---
description: See changes per version
---

# Change log

## 2.9.11

* **New** [**cloud system**](../developing-with-xlconnect/cloud-systems.md#moss)**:** Moss
* **Security fix**: logins that use Client Credentials (Visma Client Credentials, Moss) no longer send your XLConnect session along to the cloud system's login server.
* **Fix**: logging out of a named connection (for example `exact:acme`) now works.

## 2.9.10

* [**Standalone runs**](../developing-with-xlconnect/javascript/standalone-mode.md) no longer fail with _"Process … not found"_ on their first call to a cloud system.
* **Standalone runs never wait for a login window.** When a cloud system has no saved login, the run stops straight away with a message to sign in once interactively, as the Windows user the scheduled task runs under. Before, it opened a login window nobody could see and waited forever.

## 2.9.9

* **Table connections stay linked** when a sheet is renamed, also when that happens in Excel Online, on a Mac or on a computer without XLConnect. A table connection that cannot find its table now shows red in the Maps tree, and XLConnect warns after a run when data could not be written. Workbooks where a connection had already lost its table need to be relinked once.
* [**Run reports by email**](../developing-with-xlconnect/javascript/standalone-mode.md#notifying-users-of-the-outcome)**:** add `--notify failure|success|both|none` to a standalone run to email a summary of the outcome. Choose who receives it with the new **Notify** role in the portal.
* **Logins stay valid**: tokens that are refreshed automatically during a session (XLConnect, Xero, Visma, Microsoft) are now saved, so the next reconnect no longer fails.
* **Clearer error messages**: an expired session, missing permissions or a missing item now say so (with a reconnect option in Quick Open), instead of _"An exception occurred: One or more errors occurred"_.
* **Fixes**:
  * `FetchData` / `ClearData` from VBA and the ribbon Run button failed for users without the Maps pane.
  * Auto-login could be skipped, or the ribbon not updated, while Excel was still starting.
  * Two connections with the same name made the workbook fail to open.
  * Failed standalone runs now always write a `<script>.<date>ERROR.txt` next to the log, so failures are visible from the folder listing.

## 2.9.8

* [**Standalone login**](../developing-with-xlconnect/javascript/standalone-mode.md)**:** run scripts from the Windows Task Scheduler as a service account tied to one database, instead of a personal login that can expire. Set it up from Excel (**Advanced → Setup Standalone**) or from XLConnect Desktop; one computer can hold several accounts.
* **Modules load reliably**: `require(...)` now loads modules through the XLConnect server instead of directly from GitHub, which some corporate networks block.

## 2.9.7

* **Installer**: Start Menu shortcuts, more reliable upgrades, and fixes for the add-in and Desktop failing to load after an upgrade.
* **Stability**: fewer freezes and crashes, and a local log file to help diagnose errors.
* **Fix**: signing in through the browser a second time could fail.

## 2.9.6

* **Simplified ribbon**
* **Fixes**:
  * Spilling formulas (dynamic arrays)
  * Memory leak in table connections
  * Code completion for function arguments and `xlc.` methods
* Removed the old XLCTable option.

## 2.9.5

* **Code completion** in the Javascript Studio: `xlc` methods, variables in the engine, function arguments, lambdas and `for (x of list)` loops.
* **Table connections** now keep your column layout by default and only add new columns.
* **Fix**: Payhawk login dialog.

## 2.9.4

* **New** [**cloud systems**](../developing-with-xlconnect/cloud-systems.md)**:**&#x20;
  * Visma eAccounting
  * Visma Client Credentials Multi (switch between tenants)
* Now requires .NET Framework 4.8.1

## 2.9.3

* Installer fix

## 2.9.2

* **Fixes**:
  * Saving workbooks to the cloud
  * Manage Access
  * SFTP
* New installer

## 2.9.1

* [**xlc.uploadBinary**](the-xlc-object.md#xlc.uploadbinary-string-uri-string-localpath-dynamic-headers-null-string-auth-null) to upload files
* [**SFTP**](the-xlc-object.md#sftp) methods
* Ribbon buttons stay available for protected workbooks

## 2.9.0

* **SharePoint Data Lake** login (`auth = 'spdl'`), see [Data Lake (SharePoint)](../developing-with-xlconnect/data-lake-sharepoint.md)

## 2.8.11

* [**xlc.guid()**](the-xlc-object.md#xlc.guid)
* **Clear on Clear**: turn this off on a connection to keep its data when the map is cleared, for example for a table that holds settings.
* **Fix**: Nmbrs login

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

