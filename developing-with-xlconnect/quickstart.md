---
description: Build and publish your first connected workbook
---

# Quickstart

In this quickstart you build a workbook that pulls live data from a cloud system — without writing code — and publish it for your users. It assumes the add-in is [installed](../installing-the-add-in.md) and you are logged in.

## 1. Enable Developer Mode

Go to `Settings > Advanced > Developer Mode` on the XLConnect ribbon. A number of new buttons appear — see [Getting Started](getting-started.md) for a tour.

## 2. Import a prefab script

The fastest way to a working script is a prefab. On the ribbon, select `Import Script` and pick your platform — for example `Xero > Connections`. The Maps and Scripts pane opens with a ready-made script and its input and output tables.

Follow [Import Scripts](import-scripts.md) for the full walkthrough of this step and the next two.

## 3. Connect the output to the sheet

Drag the script's output table from the pane onto your worksheet. It appears as an Excel table that will derive its columns from the data.

## 4. Run it

Right-click the script and choose `Execute Script`. The first run opens a browser to log in to the cloud platform; then the data lands in your table. From here it's normal Excel: add formulas, charts and formatting on top of the live data.

## 5. Look under the hood

Right-click the script and choose `Edit javascript` to open [Javascript Studio](javascript/javascript-studio.md). Hit `F5` to run the code and watch the variables appear in the Engine pane. This is where you customize — the prefab code is yours to edit, and the [modules](../reference/modules/README.md) keep it short.

## 6. Publish

Press `Save to Cloud` to publish the workbook to the Model Library. Your users can now open it through `Open Workbook` and refresh it with `Run` — and every version you save is tracked in the `Workbook History`, with rollback if you ever need it.

## Where to go next

* [Core Concepts](core-concepts.md) — the mental model behind what you just did
* [Cloud Systems](cloud-systems.md) — all platforms you can connect to
* [Javascript](javascript/README.md) — writing your own scripts
* [Data Lake (SharePoint)](data-lake-sharepoint.md) — storing results without a database
