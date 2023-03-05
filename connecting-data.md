# Connecting a Workbook to data in the Data Lake

When using XLConnect, the most typical usecase is connecting to data that someone else will have prepared in the Data Lake. This could be interest rates that were downloaded daily, reference data like mortality tables, pricing parameters etc. Connecting your workbooks to such data is one of the core features of XLConnect.

  1.
## Importing the original format

- Open a new workbook and make sure the launch pad is visible on the left.
- Navigate to the workbook &#39;Tutorial \&gt; 2. Basics \&gt; Loan Calculator.xlsx&#39; but do **not** open the workbook. Instead just select it to display the Messages:

![](RackMultipart20200606-4-10x3hpl_html_c256385d4fb0e920.png)

You should see something like this, where the &#39;Loan Calculator&#39; workbook is selected, and the messages pane shows the three loans.

- Double click any of the messages, this should bring up this dialog:

![](RackMultipart20200606-4-10x3hpl_html_e700e233e3d0b10d.png)

- Leave the settings as they are and click &#39;Import&#39; (or press ENTER) to complete the import.

Your screen should look something like this:

![](RackMultipart20200606-4-10x3hpl_html_98bd262009c3ed52.png)

The values of the fields in the message have been put in the worksheet on the same place as they were on the &#39;Loan Calculator&#39; workbook.

- &#39;Import Original Format&#39; imports a message by putting the cells in the same layout as they have in the workbook that produced them.

- Double click any of the other loans in the messages pane to see that not just the values have been imported, but connections have been made to load that message type into workbooks.

- Importing massages not only imports the data of a particular message, but creates connections to load that type of message.

- If you would like you can commit this workbook to be able to se it again. Else just close it to start for the next exercise.

## Importing the Message format

- Create a fresh workbook by starting Excel or pressing CTRL-N to create a new workbook.
- Navigate to the workbook &#39;Tutorial \&gt; 2. Basics \&gt; Loan Calculator.xlsx&#39; again and do **not** open the workbook. Instead just select it to display the Messages.
- Double click any of the messages (loans) to bring up the &#39;Import Message&#39; dialog again
- This time select &#39;Message&#39; format and press &#39;Import&#39;. You should now see this:

![](RackMultipart20200606-4-10x3hpl_html_fe094c1735108a2f.png)

These values form the message were not imported in the layout of the original workbook, but they are just imported in the order they appear in the message. Notice how the arrays (Period – RepaymentActual) are now rows instead of columns.

- Import Message Format imports a MessageType and lays out the cells in the order they appear in the message.
- XLConnect imports arrays as rows by default, these can be made columns. See paragraphs 7.4 and 7.5 to see how.

## Importing Map Only

- Open up a new workbook and navigate to the loans messages again like in the previous two examples
- Double click to bring up the Import dialog, this time choose &#39;Map only&#39; and click &#39;Import&#39;

You should now notice no values have been added into the worksheet, only that the map has appeared on the right.

- Double click any of the other messages to see their values are not put into the sheet
- Keep the workbook open for the next exercise

## Drag &amp; Drop

- In map1, open the object &#39;Input&#39; and drag the connection &#39;Loan&#39; to cell B2:

![](RackMultipart20200606-4-10x3hpl_html_98bc14db9dc29bbe.png) ![](RackMultipart20200606-4-10x3hpl_html_bff66c5cc67a528d.png) ![](RackMultipart20200606-4-10x3hpl_html_4c9f2457c80a82aa.png)

Dragging a conection onto the workbook connects that cell.

- Now drag the ArrayConnection &#39;Calculations/Principal&#39; onto the workbook:

![](RackMultipart20200606-4-10x3hpl_html_4cfd0daead4c37f3.png)

![](RackMultipart20200606-4-10x3hpl_html_f7cc57efeda4e08.png)

This will connect the array.

- Now hold down the CTRL key and drag ArrayConnection &#39;Calculation/Month&#39; onto the workbook.

![](RackMultipart20200606-4-10x3hpl_html_dd4fb5f9fc0db17d.png) ![](RackMultipart20200606-4-10x3hpl_html_993a6348d399f4d1.png)

Now the array is inserted not as a row but as a column

- Holding down the CTRL key when dragging ArrayConnections pivots them into row columns. This only works for ArrayConnections, RangeConnection and TableConnection remain in their original orientation.

- Drag &#39;Calculations/InterestPayment&#39; next to column period (as a columns by holding down CTRL)
- Drag &#39;Calcualitons/PrepaymentActual&#39; in between these two columns
  - Hold down CTRL at the beginning of the drag
  - Hold down SHIFT to insert the cells between them

![](RackMultipart20200606-4-10x3hpl_html_215c441ae80fa94a.png) ![](RackMultipart20200606-4-10x3hpl_html_e9e4dfd01d2053b4.png) ![](RackMultipart20200606-4-10x3hpl_html_a8e3cae97a97d116.png)

- Holding down the SHIFT key on any connection inserts the cells between other cells. This can be combined with the CTRL key for ArrayConnections to insert them as columns.

## Updating connections

So far we&#39;ve discussed creating connecting cells from the workbook, or dragging them from the Map. There are two more ways to changes connections: updating connections and disconnecting

\&lt;\&lt; todo \&gt;\&gt;

## Importing a message format multiple times

## Advanced topics

# Writing Data to the Data Lake with a Workbook

In the previous chapter we connected workbooks to data others had prepared for us in the Data Lake. In this chapter we will show how to create such data from workbooks.

## Connecting cells

- Open the workbook &#39;Tutorials \&gt; 4. Creating Connections \&gt; Loan Calculator No Connections.xlsx&#39;
- Select cell C2 (with value 10.000)
- Click the &#39;Connect&#39; button on the ribbon (Click the top half of the button with the icon, not the bottom half that opens up a dropdown that will be explained later)
 ![](RackMultipart20200606-4-10x3hpl_html_f39a979d97bb22f5.gif)

![](RackMultipart20200606-4-10x3hpl_html_f1cb15821b054b0c.png)

Now the &#39;Maps&amp; Messages&#39; task pane should appear showing this:

![](RackMultipart20200606-4-10x3hpl_html_b6d4c34b47fd2c7f.png)

This shows that new Map has been created called &#39;Map1&#39; with one connection called &#39;Loan&#39;.

This name was taken from the cell to the left of C2 because that&#39;s usually where the label for a cell will be.


## Changing connection names

The name Map1 isn&#39;t very helpful, so we want to

- Select the map called &#39;Map1&#39; and hit F2, this will put the Map in edit mode. Change the name to &#39;Loan&#39; and hit enter. Hitting enter will cycle through the connections to keep changing the names.
- As we are content with the name right now hit ESC to exit edit mode.

- F2 puts the Maps&amp;Messages control into Edit Mode, ESC goes back out.

Your Maps&amp;Messages pane should now look like this:

![](RackMultipart20200606-4-10x3hpl_html_a9006e05a220a4ec.png)

- Now select range C3:C4 (so the two cells that contain 36, and 2.40%) and click on the bottom half of the connect button to open the dropdown, and select Property.
 ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif)


![](RackMultipart20200606-4-10x3hpl_html_8a564a9112fd957.png) ![](RackMultipart20200606-4-10x3hpl_html_bb8cde19d9c0a790.png) ![](RackMultipart20200606-4-10x3hpl_html_570d82a3e2d05790.png)

You should now see two connections added to the Map called &#39;Loan&#39;

- When you select multiple cells and choose to add as Property Connection, each individual cell will be added as separate connection saving quite some of clicking around.

You may wonder where it got the names from. XLConnect looks to the left of a newly connected cell to see if there is a word there, as that is the most likely place for a meaningful label to be. If it is not as desired you can change the connection name with F2. As the names are fine we will leave these as is.

- XLConnect looks to the left of newly connected cells for an initial name for a new connection. You can always edit it with F2.

Now we have our first connections on this workbook, we want to save it somewhere. As you cannot save to the tutorials database we will save the workbook in your personal folder.

- On the ribbon, hit the Commit Workbook button, and save the workbook to your personal folder (See 5.3 Committing the workbook to recall how).
- Close Excel

## Types of connections

XLConnect has different types of connections:

1. PropertyConnections which contain a single value
2. ArrayConnections that contain an array of values
3. RangeConnections that contain a range of multiple rows and columns. It will be stored as an array of arrays.
4. TableConnections that contain a datatable
5. ObjectConnection which is a collection of other connections. These allow you to structure larger messages into &#39;chapters&#39; or create specific layouts required to communicate with other systems that expect a specific layout.

- Open Excel again, then open the loan calculator from your Personal folder by any to two ways:

  1. In the Launch Pad, open your personal database and double-click the file
  2. From the ribbon, select quick open \&gt; your personal db \&gt; the workbook.

- Select range B10:B70 (those are the period numbers 0-60) and click on the top half of the Connect button.

![](RackMultipart20200606-4-10x3hpl_html_91ee2d13ead3c594.png)

You will now see that a Connection called &#39;Month&#39; is added. If you look at the icon, it is not a single blue square but a blue row. This is the icon for an ArrayConnection.

Also note that in the case of a column, the label was taken from the cell above, instead of to the left.

We will now connect the remaining cells

- Select range C10:E70, those are the values for the three columns Principal, Interest and Repayment
- Click
  - bottom half of the connect button
  - then the \&gt; indicator to the right of Array
  - Split area into columns

![](RackMultipart20200606-4-10x3hpl_html_dfed49db78e03534.png) ![](RackMultipart20200606-4-10x3hpl_html_7f84e436ef4e0880.png)

This will have added the three columns as array connections. This splitting can also be done to connect larger areas of cells in which case this will save quite a bit of repetitive selecting and clicking.

What you will notice is that there are two connection named Interest that are colored red.

- hover above one of them to read the validation message: &#39;The name Interest is used more than once on this level. Change one, or move to another level by inserting an object&#39;

That message is pretty self explanatory. If you try to save the workbook now you will get the message

![](RackMultipart20200606-4-10x3hpl_html_fb76612a68173f17.png)

- Press F2 and change the second connection into &#39;InterestPerPeriod&#39;, press ENTER and then ESC.

The map should now pass validation:

![](RackMultipart20200606-4-10x3hpl_html_b1cf0a1f57cabbcf.png)

- Hit CTRL-S, this should bring up the Commit Workbook dialog
- Press ENTER to commit changes

- CTRL-S, ENTER is way to commit your workbook changes in under a second with just two keystrokes, although it will help a third party and your future self if you add comments on the changes you made.

## Guessing connection type

In 6.1, when we selected a single cell and pressed the connect button (top half, without specifying the type) it created a PropertyConnection. When we selected a Single column of cells it created an ArrayConnection.

By default, when you press the top half of the Connect button (or hit CTRL-SHIFT-C which is the shortcut for the Connect button), XLConnect will &#39;guess&#39; the connection type best suited to your selection by using there rules:

  - Is it a single cell: PropertyConnection
  - Is it a single column or row: ArrayConnection
  - Is a range with multiple rows and columns: RangeConnection
  - Is it a Table: TableConnection
  - When a multiple selection is made (by holding down CTRL and selecting other cells, the rules are applied to every section).

- When you have selected multiple cells and select property, every cell is added a separate property.

There are quite a few ways in which XLConnect can save effort to connect new worksheets, especially when they contain a lot of cells knowing the behavior can save a lot of work.

## Different ways of creating connections

There are multiple places in the GUI where you can

- CTRL-C
- Ribbon
- Task pane
- Context Menu (right click)