---
description: Jumpstart your development by importing prefab scripts
---

# Import Scripts

The easiest way to get started with XLConnect development is with Prefab Scripts, that you can import from the Ribbon. For this example we will use Xero, for other platforms it works similarly.&#x20;

On the Ribbon, select `Import Script > Xero > Connections`.&#x20;

<figure><img src="../.gitbook/assets/image (21).png" alt=""><figcaption><p>Insert a Xero Connections Script</p></figcaption></figure>

On the right hand side of your screen the Maps and Scripts pane will open with a script called 'Connections', with under it two maps Input and Output, and under output a table again called Connections.&#x20;

![](<../.gitbook/assets/image (24).png>)

The Connections table is grey, indicating it is not connected to workbook cells. To connect it, click and hold the table, to drag it onto your workbook. This should create a table with two generic columns called Column1 and Column2, and a message telling you it will get it's columns from the data:&#x20;

![](<../.gitbook/assets/image (26).png>)

So let's execute the script. To do that, right click the script, then choose `Execute Script`:

![](<../.gitbook/assets/image (27).png>)

This will execute the script, this will call into the Xero API to pull the connection so it will probably open up a browser to log in to Xero.&#x20;

![](<../.gitbook/assets/image (30).png>)![](<../.gitbook/assets/image (31).png>)

Complete the steps, then it will pull the data and display it in the table:&#x20;

<figure><img src="../.gitbook/assets/image (29).png" alt=""><figcaption><p>Table with list of Companies from Xero</p></figcaption></figure>

The companies in this example have been altered, you should recognize the companies on Xero that you have access to. To get access to more companies, go though the login process again.&#x20;

Now it's time to select the company we want to pull the data for, to do that mark the column `Pull` with an x.&#x20;

![](<../.gitbook/assets/image (34).png>)

Now we have a list of Zero Companies, but no data yet. Let's pull the chart of accounts for a company. To do this, import the Accounts script:&#x20;

![](<../.gitbook/assets/image (32).png>)![](<../.gitbook/assets/image (33).png>)

This imported a second script 'Accounts', with one input table `Connections` and an Output table `Accounts`. The Connections table is white, inidcating it is connected. That's because it the Excel table created in the step above is also called Connections. XLConnect will automatically connect a table connection when it finds a table with the same name.&#x20;

To connect the Accounts table to the workbook, drag it onto the worksheet. You will get the same message a earlier that it needs to pull data to derive the columns. Your workbook should look like this:&#x20;

<figure><img src="../.gitbook/assets/image (36).png" alt=""><figcaption><p>Empty Accounts table below Connections table</p></figcaption></figure>

To expand the table, we'll do the same as with the first table, execute the `Accounts` script:&#x20;

![](<../.gitbook/assets/image (37).png>)

Now XLConnect should grab your Chart of Accounts data from Xero and display it the table:&#x20;

<figure><img src="../.gitbook/assets/image (38).png" alt=""><figcaption><p>Chart of Accounts for Company 3</p></figcaption></figure>

That is how easy it is to grab data from Xero by Drag & Drop, without even writing a single line of code.&#x20;
