---
description: How to perform daily tasks in XLConnect
---

# Using XLConnect

The daily tasks, one by one. If you're brand new, do the [Quickstart](quickstart.md) first.

<!-- TODO: add screenshots for each section -->

## Logging in and out

Press `Login` on the XLConnect ribbon. Your browser verifies your identity (via your Microsoft or Xero account) and returns you to Excel. XLConnect shares its login with the web app — if you're logged in there, Excel picks it up.

To switch accounts: log out in the web app (`Settings > Jump to Webapp`), then press `Login` in Excel again with the other account.

## Opening workbooks

There are two ways to open a workbook from the cloud:

* **Open Workbook** on the ribbon drops down the databases you have access to. Hover to browse folders; a tooltip shows each workbook's description and who last edited it. Click to open.
* **Start Menu** shows the same content organized as a launcher, when your organisation has one configured.

Because workbooks come from the Model Library, everyone always opens the intended version — no more copies mailed around.

## Refreshing data

* **Run** executes the workbook's scripts: connect to the cloud, pull fresh data, land it in the sheet. A progress bar shows what's happening.
* **Clear Data** empties the loaded data for a clean start — for example before selecting another company or period and running again.

## Cloud system connections

The first time a workbook pulls from a platform like Xero, a browser opens to log in there and choose what to share. The connection (token) is then stored securely.

To disconnect or switch to a different environment, go to `Settings > Advanced > Manage Tokens`, right-click the connection and choose `Delete`. The next `Run` will take you through the login again.

## Invites

When someone gives you access to a database, you'll get an invite. `Settings > Check for Invites...` lets you accept pending invites if you missed the email.

## Getting help

The `Support` group on the ribbon lets you create a ticket (`Create ticket...`) and see the status of earlier ones (`My tickets...`). See [Troubleshooting](troubleshooting.md) for the common issues and their fixes.
