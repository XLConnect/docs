# The Bigger Picture

This chapter will explain the core ideas behind XLConnect. It does not contain any detailed steps how to performs certain tasks, rather it explains why you would want to do them in the first place. Later chapters will teach how to do them.

XLConnect revolves around four pillars:

1. Separating data from logic in Excel
2. Strict version control of workbooks and data
3. Workflow, Enterprise models and process automation
4. External connections (API&#39;s, Databases)


## Separating data from workbooks
![](2020-06-06-20-24-20.png)
The first core feature of XLConnect is to separate data from the logic. In classic Excel, a workbook contains formulas, user interface markup (like labels, lines and charts), VBA and data. For instance when you have created a simple loan calculator and want to evaluate three alternative offers, you would create three copies of that workbook (or at least create copies of the formulas). Each loan would have its own copy of not only the data but also the formulas.

This has some drawbacks, especially when you need to prove that the results are correct like in current financial institutions. When dthe workbook is copied and changed, these changes need to be four-eyed, both the data and formulas. This can quickly add up when you have thousands of loans.

In &#39;normal&#39; programming (C#, Java, Python, etc.) the logic is separated from the data. To calculate three loans, you would have a single piece of code (a function) that you then call with three different sets of inputs. The advantage of doing it that way if that you have to maintain and validate only one copy of the code that you reuse for each loan.

XLConnect allows you to do the same by exchanging data in and out of workbooks. This way a workbook that is used many times only needs to be checked once, and checking the data is much less work. XLConnect also provides support for checking and approving workbooks and data. When the correctness of data from Excel needs to be evidenced to (external) auditors, this functionality will save a significant amount of time while delivering clear guarantees.

### Workbooks, Maps and Messages

In the loan calculator example above, in classic Excel you would have three copies of the workbook. In XLConnect you can have a single copy of the workbook (with the logic and user interface) and three separate copies of the data called &#39;Messages&#39;. The picture below shows a schematic representation

![](RackMultipart20200606-4-10x3hpl_html_2384cac91c13ca38.gif)

Instead of storing the details of each loan in a copy of the workbook, the data is stored outside of the workbook in so called messages.

There are quite a few advantages of separating the data from the workbook, to name but a few:

- Reduced effort of checking whether formulas are correct
- A workbook can be used without creating a copy of it, which leads to many copies of copies of workbooks.
- Versioning is handled separately which reduces effort and increases clarity
- Ability to change a workbook and recalculate all messages (instead of either updating the formulas in every copy or copying the new workbook and copying over the data)
- Messages in the Data Lake can be searched and aggregated (compared to workbooks in local disks, mailboxes and network shares).
- Automatic data handling eliminates fat thumbs finger risk. XLConnect tracks changes to a workbook to guarantee only intended cells were changed, on top of tools to lock workbooks.

Messages are swapped in and out of workbooks through Maps that are groups of Connections. Each connection points to cell or range of cells in the workbook and gives those a name. When a message is extracted from the workbook, every connection extracts the values in its cells and saves under them as its name, and the all these name-value pairs are stored in a Message. This way a map produces a message with all the values in it.

![](RackMultipart20200606-4-10x3hpl_html_8ee9a468d025383d.gif)

Here is how that looks in XLConnect, the &#39;Loan Calculator.xlsx&#39; can be found in folder &#39;2. Basics&#39; of the Tutorial database.

![](RackMultipart20200606-4-10x3hpl_html_c86686ed61219811.gif) ![](RackMultipart20200606-4-10x3hpl_html_ee90c64248267677.gif) ![](RackMultipart20200606-4-10x3hpl_html_52cd4936f1f2e2a9.gif) ![](RackMultipart20200606-4-10x3hpl_html_e060abc17eb2b2a.gif) ![](RackMultipart20200606-4-10x3hpl_html_e060abc17eb2b2a.gif) ![](RackMultipart20200606-4-10x3hpl_html_e060abc17eb2b2a.gif) ![](RackMultipart20200606-4-10x3hpl_html_e060abc17eb2b2a.gif) ![](RackMultipart20200606-4-10x3hpl_html_89ed9cd5d6d5b44.gif)

Messages

Map

Workbook

 ![](RackMultipart20200606-4-10x3hpl_html_9c31c5b4446939f0.png)

The map contains a number of connections that point to a range of cells in the workbook and gave that a name. The bottom right corner shows the messages that can be loaded into the workbook like we did in the &#39;Hello world&#39; example.

The picture below shows the contents of these messages. We can also inspect the message that was created, this should have the same value screen can be opened by right-clicking a message and then selecting &#39;Trace lineage…&#39;. This will bring up the values in a message and everything else you need to inspect that message.

![](RackMultipart20200606-4-10x3hpl_html_f54625ad05c75216.png)

The picture below shows the contents of these messages. The names defined in the map on the left match those in the message and the values correspond to the values in the workbook above.

![](RackMultipart20200606-4-10x3hpl_html_484c2b0fc5f41e1e.gif) ![](RackMultipart20200606-4-10x3hpl_html_f24b18b3f3795f33.gif) ![](RackMultipart20200606-4-10x3hpl_html_f24b18b3f3795f33.gif) ![](RackMultipart20200606-4-10x3hpl_html_ee90c64248267677.gif) ![](RackMultipart20200606-4-10x3hpl_html_52cd4936f1f2e2a9.gif)

Message

Map

 ![](RackMultipart20200606-4-10x3hpl_html_193fa2413864d681.png) ![](RackMultipart20200606-4-10x3hpl_html_fd0bc93a3f635b3c.png)

A map is a go-between between the workbook and the message, it connects the names in the message (format) with cells in the workbook. When a message is extracted the map takes the values from the workbook and saves them in the message. When a message is loaded a map will direct the named values to the connected cells.

Maps are part of the workbook and are saved with it. The messages that are extracted are stored separately from the workbook as there can be many from a single workbook.

This allows for a much more intuitive way of exchanging data then using the traditional approach with VBA and a (relational) database. If you&#39;ve ever tried that you will quickly appreciate how XLConnect will save you technical knowledge, effort and room for bugs.

  1.
## Version control of workbooks and data

The second core feature is strict control over the versions of workbooks and data. Versioning is a bit of a problem in Excel because it depends on the file system. Suppose Alice created a spreadsheet called MySheet.xlsx. She then emails this spreadsheet to Bob. Now both have identical copies of the workbook with the same name. Now Bob makes some changes and presses save. Now both have different versions of a workbook that are both called MySheet.xslx.

This quickly becomes complicated when collaborating in teams. It requires discipline in sticking to naming conventions and making sure you always open the latest version, yet most will have personal experiences where this fell apart under time pressure. Also we will all have seen spreadsheets called &#39;MySheet v3.2 Final \_ Comments Joe.xlsx&#39;. The label &#39;final&#39; in Excel names rarely means what it says in the dictionary.

XLConnect allows for a much more strict versioning of both workbooks and data saved in messages.

### Versioning Workbooks

Workbooks can be stored on the platform using the Commit command. If you are familiar with source control tools like GIT of SVN you will recognize that word. Commit means saving a particular version of code in such a way that you can always get it back later. XLConnect does the same for your Excel sheets. When you commit a workbook to XLConnect it will do a few things:

- Give that workbook version a unique identifier
- Get an MD5 checksum of that workbook file
- Store which version this version was derived from
- List the changes compared to that previous version in both formulas and VBA
- Store metadata like user, date and comments

All these details are collected and stored automatically (and completely) saving excel developers a lot of laborious work.

### Versioning Data

When producing data with Excel, using correct inputs is as important as using the right transformations (workbooks). XLConnect puts data in the Database that provides the same controls as with workbooks.

For every message, metadata is stored about

- who produced it
- when
- from which workbook
- which messages were used as inputs to that workbook
- who approved it

All this logged automatically. With a relational database like Access or SQL Server you would have to program these features into every table that is created which is a lot of repetitive work. With XLConnect these features come out of the box allowing you to quickly build applications that already tick all the compliance boxes without you having to worry about it.

#### Data Lineage

Every message keeps track of which workbook and other messages were used to produce it. For these messages the same is stored all the way to the first input. XLConnect allows you to navigate this chain intuitively to always understand which data was used in a particular report. The same goes for the workbooks, you can drill through all the versions to understand which changes were made and when. These two lineages combine to paint the complete picture of how a particular message was produced. This is very useful for audit purposes.

#### Mutable vs Immutable

When you approve a particular piece of information, you do not want people to be able to change it afterwards. The same goes for drilling through the lineage described above, that only reproduces the correct picture when data is not changed. Therefore Messages (MessageVersions to be precise) are immutable. They can written only once, after that they cannot be changed.

In practical applications however one needs to change data for all sorts of reasons. When that is required, you write a new message version to the MessageHeader, which effectively is a stack of MessageVerions. If you look at a message header and don&#39;t specify anything, you will see the last written message on top of the stack. But older versions are always available. This way the data appears to be mutable for practical applications, but the audit and lineage functions will also work. This same principle applies to workbooks. For more detail how headers and versions are used please refer to chapter 11 – Headers and Versions.

  1.
## Workflow, Enterprise models and process automation

The third feature is creating and automating workflows, and modelling larger structures out of connected messages. \&lt;\&lt; More on this later, most features not completely ready yet \&gt;\&gt;

  1.
## External Connections (API&#39;s, Databases)

The fourth pillar of XLConnect is the ability to communicate with external systems through REST API&#39;s or ODBC connections.

XLConnect itself is built upon REST API&#39;s that communicate JSON, this way of building cloud API&#39;s is a widely used standard for building API&#39;s. This makes it very natural to connect to REST API&#39;s created by others to pull in data for your workflows.

For instance you can pull in data about [stocks](https://www.alphavantage.co/query?function=GLOBAL_QUOTE&amp;symbol=MSFT&amp;apikey=demo), interest rates, [weather](https://samples.openweathermap.org/data/2.5/forecast/daily?q=M%C3%BCnchen,DE&amp;appid=b6907d289e10d714a6e88b30761fae22) or [credit scores](https://anypoint.mulesoft.com/mocking/api/v1/sources/exchange/assets/5cc8ad1a-2f92-447f-9c4c-a48e6560cf4f/company-api-v1.0.0/1.0.20/m/nl/companies/302459/scores/creditScores?client_id=5cc8ad1a2f92447f9c9ca48e6560cf4g&amp;client_secret=7064bba96a944z90B64Cd8025Ae30A7c), to name but a few. The amount of data that is available through REST API&#39;s in enormous and can now be used in workflows with a few clicks.

Alternatively, data can also be written to systems that support REST API&#39;s, this includes systems like SAP, Navision, Salesforce. This connection also works two ways: XLConnect can call the systems, but these systems can also call XLConnect to retrieve data or call workbooks. This way XLConnect can be used to quickly build workflows that can be integrated into these systems. Excel is a much more productive tool to create computations in than those platforms, combining the two combines the best of both worlds.

If on premise systems do not yet provide REST API&#39;s, XLConnect also supports querying databases with parameterized SQL statements through ODBC connections. This can pull in data from these systems for reporting or other purposes.
