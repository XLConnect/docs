---
description: >-
  What XLConnect is, the problems it solves and the core ideas behind it
---

# What is XLConnect?

XLConnect is a platform that connects Excel to the cloud. It extends Excel with the ability to read and write data from APIs (like Xero), store data in a data lake (like SharePoint) and automate workbooks — all scripted in Javascript, from inside Excel.

Excel is uniquely productive for financial modelling and reporting: the person who understands the business can build the model themselves, without translating requirements to a development team. But used in teams over longer periods, plain Excel shows weak points: manual work, versions mixed up in email, copy-paste that obscures where numbers came from.

XLConnect keeps the strengths and fixes the weak points:

* **Data management** — read and write data between workbooks, cloud APIs and the data lake, instead of copy-pasting it around.
* **Automation** — run workbooks unattended, so hundreds of workbooks can be refreshed with new data automatically.
* **Version control** — workbooks live in the Model Library, which tracks every version and lets you roll back. Everyone can see which version they are using.
* **Governance** — integrated approvals for data and workbooks make results reproducible and easy to report on.

## The core ideas

Three concepts carry the whole platform. Understand these and the rest of the documentation will fall into place:

* **Scripts** — Javascript that runs inside your workbook. Scripts read and write cells, ranges and tables, call cloud APIs and move data in and out of the data lake. You can [import ready-made scripts](developing-with-xlconnect/import-scripts.md) or write your own in [Javascript Studio](developing-with-xlconnect/javascript/javascript-studio.md).
* **Messages** — the JSON documents that data travels in. A message can hold the inputs of a model, its results, or reference data shared between workbooks.
* **The data lake** — where messages are stored, without you having to design a database first. Save a message from one workbook, load it in the next, and you have a workflow.

## Where to go next

* If you will be **using** workbooks that someone else built: [Using XLConnect](using-xlconnect/getting-started.md)
* If you will be **building** connected workbooks: [Developing with XLConnect](developing-with-xlconnect/getting-started.md)
