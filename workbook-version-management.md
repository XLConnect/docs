# Workbook Version Management

A second core feature of XLConnect is very strict control over the versions of workbooks. This part of the manual will show how to make and review changes to a workbook, commit versions to the server and create overview of all the different versions and changes between them.


## Tracked Changes Formulas

When you change cells in a workbook, XLConnect will automatically record these changes.

- Click on the button &#39;Formula Changes&#39; on the ribbon to show/hide the &#39;Formula Changes task pane&#39; on the bottom of the screen.

![](RackMultipart20200606-4-10x3hpl_html_5f37f2d2ce4979a3.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif) ![](RackMultipart20200606-4-10x3hpl_html_2f1cee0fa4fd0858.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif)


 ![](RackMultipart20200606-4-10x3hpl_html_b671e77a9d74d9f5.png)

- Start typing and making changes to the workbook. You will see these changes appear in the &#39;Formula Changes&#39; task pane.
- Select rows in the list of changes to see those cells activated in the worksheet to help you see where they occurred. Also the details of the before and after situation is shown in the right half of the screen

The right half of the screen shows the before and after situation for that change. The text colors indicate what happened in that change:

- Red has disappeared
- Green is new
- White means unaffected


## Tracked Changes VBA

XLConnect also tracks changes in VBA which is helpful as these are often hard to detect and review.

- Hit ALT-F11 to open up the VBA editor ![](RackMultipart20200606-4-10x3hpl_html_7350c341c0917262.gif)
Sub SayHello()msbox &quot;Hello!&quot;End Sub
- Double click ThisWorkbook to open up the code for the workbook
- Paste in the snippet to the right:
- Return to the workbook by hitting ALT-F11 again
- Click on the ribbon button &#39;VBA changes&#39;
- You should now see this screen: ![](RackMultipart20200606-4-10x3hpl_html_baa763b980c7f38.png)

The colors have the same meaning as with the formula changes.

- The VBA Changes dialog tracks the changes you make to VBA projects. This information is stored in when you commit the workbook for others to review.


## Committing a workbook

Now we have made changes to the formulas and VBA of our workbook, we want to &#39;commit&#39; it to the server.

**Note** : Committing is different from Saving a workbook. Saving is still possible but to use the XLConnect functionality you need to Commit a workbook. A workbook opened from disk instead of XLConnect will not contain the maps or record changes. Only workbooks committed to and opened from XLConnect have those.

- Click on the Commit button on the ribbon, this will bring up the &#39;Commit dialog&#39;

![](RackMultipart20200606-4-10x3hpl_html_a6b9464b9b20f4fc.png)

- Give the workbook a meaningful name (anything starting with &#39;Book&#39; is not allowed to prevent names like Book1 and Book2 making it into the model library). Take a second to give your workbook a name that means something.
- Select a location by clicking the small … button next to the location field. For now select your personal database.

Click outside the dropdown to close it again.

- Short description is what people will see as a tooltip before opening a workbook, this is helpful in describing what the workbook does.
- The version comments allow you to add comments to this particular version you commit. This would contain a few short sentences to help yourself or someone else to quickly understand why this version was created.
- On the list of changes, columns &#39;Work item&#39; and &#39;Comment&#39; are editable, they can be used to add comments on specific changes.
- Then hit commit (or press return) to complete the commit.
- When you return you will see that the changes have been cleared because changes are always compared to the previous version that you just committed. The current workbook is identical to that. The same goes for VBA.

