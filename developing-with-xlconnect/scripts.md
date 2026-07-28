# Scripts

Scripts are the working horses of a connected workbook: javascript that reads inputs from the workbook, talks to APIs and the data lake, and writes outputs back to cells and tables. Users run them with the `Run` button — developers build and edit them as described here.

## Adding scripts

With [Developer Mode](getting-started.md) enabled, scripts can be added to a workbook in several ways:

1. **Import a prefab script** — `Import Script` on the ribbon gives you ready-made, working scripts for cloud platforms like Xero. This is the recommended starting point; see [Import Scripts](import-scripts.md).
2. **Write one from scratch** — open the `Maps and Scripts` taskpane and add a script, then build it up in [Javascript Studio](javascript/javascript-studio.md).

Scripts appear in the `Maps and Scripts` taskpane with their input and output connections. Drag outputs onto the worksheet to connect them to cells and tables, or use `Connect Cells` on the ribbon. Scripts can be re-ordered by drag & drop — they run in the order shown.

## Editing and running

* Right-click a script and choose `Edit javascript` to open it in [Javascript Studio](javascript/javascript-studio.md), where `F5` runs the code and shows every variable.
* Right-click and choose `Execute Script` to run a single script in place.
* The `Run` button on the ribbon runs the workbook's scripts — this is what your users will press.

## Running scripts outside Excel

Scripts don't need Excel to run:

* **XLConnect Desktop** is a standalone Javascript workbench — run `XLConnect.Desktop.exe` on a machine without Excel and you get the same Javascript Studio experience.
* **[Standalone mode](javascript/standalone-mode.md)** runs a script from the command line or the Windows Task Scheduler, optionally under a service account — this is how you automate daily data pulls and refresh workbooks unattended.
