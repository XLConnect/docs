# Headers and Versions

On the one hand we want workbook version to stay the same, on the other we need to change and move this to progress. These requirements are slightly paradoxical. XLConnect solves this by separating headers and versions:

- A version is a write once copy of a workbook or message. Once written it gets a unique identifier (key) and cannot be changed afterwards. Versions are write once, immutable and are typically not deleted, unless to purge the system of obsolete data. When you have a MessageVersionId or WorkbookVersionId, these will always return exactly the same value to the last byte. You cannot change it afterwards, you have to make another (derived) version. This is a very helpful guarantee for all kinds of purposes like making sure everyone is on the same page, the full lineage of every item, backups and caching, reproduce a calculation or report. All these things become much easier when you can trust on this rule to be guarded. Versions does exactly that. But of course to practically use any system you need to change data and move it around. That&#39;s what the headers do.
- A header is effectively a stack of versions, along with other things that can change over time. When you &#39;save&#39; a workbook or message, you write a message to that header. If it doesn&#39;t exist it is created. This places the new version on top of the stack of versions. The stack contains all the versions that were written to the header in the order they were written. When you ask for the value of the header you will get the top version. But every previous version is also available. Besides the stack, headers have a key (or name), a location and meta data about who changed the versions. This way can have our cake and eat it too. Move things around without changing a single byte in it.

Versions and headers are for Workbooks and Messages. The following diagram shows their relations.

![](RackMultipart20200606-4-10x3hpl_html_5767032f0edb1066.gif)

Tis diagram above shows the different types and how they connect to each other. The letters and number indicate 1..n relationships, where often when these counts match up for different classes. For instance, when you have C (say 20) connections in a map, a message created from that map will have 20 name-value pairs in it. That&#39;s why the 1..C cardinality is shown on both the Map-Connection relation and the MessageVersion-NameValuePair.

This is a short verbal description of the picture above:

- A WorkbookHeader has a stack of multiple WorkbookVersions
- Every WorkbookVersion is derived from one WorkbookVersion (which is null for new Workbooks)
- A WorkbookVersion has zero or more Maps
- Every Map has zero or more Connections (although a Map without Connections is useless)
- Every WorkbookVersion has a single Excel Workbook (a binary image of that workbook file that is stored precisely and cannot change so it can be reproduced indefinitely.
- Every connection points to a single cell or a range of cells in the workbook. (This range can actually be a compound range eg (E6:F6,H6:H7,H3,I3) fo ArrayConnection, which is considered to be a single (compound) range)
- A Map Produces a MessageVersion, which is written to a MessageHeader.
- A MessageHeader has a stack of multiple MessageVersions
- A MessageVersion has multiple Name-Value-Pairs that were extracted from the connections. For one message, the cardinality Connection:Name-Value-Pair is 1..1, but when multiple (MV) MessageVersions are extracted from a Map, the cardinality between Connection and name-value-pairs becomes 1..MV.
