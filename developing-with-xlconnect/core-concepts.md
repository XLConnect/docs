---
description: 'The mental model: scripts move messages between Excel, APIs and the data lake'
---

# Core Concepts

<!-- TODO: expand each section. This page establishes the vocabulary the rest of the developer docs rely on. -->

## Scripts

Javascript that lives in the workbook and runs inside Excel. Scripts read and write cells, ranges and tables, call cloud APIs and exchange messages with the data lake.

## Messages

The JSON documents data travels in. <!-- TODO: message types, headers, keys, versions -->

## The data lake

Where your data is stored — no database design required. The recommended option is the [SharePoint Data Lake](data-lake-sharepoint.md), which stores JSON files on your own Office 365 subscription with the easiest syntax. XLConnect also offers its [own data lake](data-lake-xlc.md), used by existing workbooks. <!-- TODO: locations, folders -->

## Cloud systems

External APIs such as Xero that scripts authenticate to and exchange data with. See [Cloud Systems](cloud-systems.md).

## How they fit together

<!-- TODO: one diagram: Excel workbook <-> script <-> (cloud APIs | data lake) -->
