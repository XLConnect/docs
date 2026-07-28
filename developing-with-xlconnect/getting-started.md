---
description: How to get started with developing XLConnect workbooks
---

# Getting Started

## Enable Developer Mode

The first thing to do to develop connected workbooks with XLConnect is to enable `Developer Mode`. When you do **not** have this enabled, your ribbon will looks like this:&#x20;

<figure><img src="../.gitbook/assets/image (15).png" alt=""><figcaption><p>XLConnect ribbon in User Mode</p></figcaption></figure>

To enable developer options, go to `Settings > Advanced > Developer Mode`

<figure><img src="../.gitbook/assets/image (19).png" alt=""><figcaption><p>Developer Mode Menu Option</p></figcaption></figure>

If you have the developer option enabled for your account, your ribbon should now look like this (notice a number of new buttons appearing):&#x20;

<figure><img src="../.gitbook/assets/image (45).png" alt=""><figcaption><p>XLConnect Ribbon in Developer Mode</p></figcaption></figure>

## Finding your way around

Here is a quick rundown of the new buttons on your ribbon:&#x20;

* **Taskpanes**
  * **Maps and Scripts** opens a taskpane on the right where you can see and edit the [Maps ](maps.md)and [Scripts ](javascript/javascript-studio.md)in the workbook.&#x20;
* **Develop**
  * **Import Prefab** let's you [import ready made scripts](import-scripts.md) to connect to your cloud platform by dragging and dropping elements, without having to write any code
  * **Javascript Studio** is the [javascript editor](scripts.md) that comes with XLConnect, that allows you to write code and see the results.&#x20;
  * **Connect Cells** lets you [add workbook cells](maps.md#adding-cells-ranges-and-tables-to-a-map), ranges and tables to maps and scripts.&#x20;
  * **Save to Cloud** saves your workbook to the database, so users can open it through Open Workbook.&#x20;
  * **Other**
    * **Sheet Sizes** lets you see the Used Ranges in your workbook, to find cells that appear empty but are 'used', bloating your workbook size without adding any value.&#x20;
    * **Analyze Formulas** lets you drill through formulas in a tree format instead of the normal arrows.&#x20;
  * **Audit**
    * **Workbook history** shows the all the versions of the workbook that have been save to the cloud, allowing you to rollback to a sepcific version if required.&#x20;
    * **Model Library** shows all workbooks on the cloud with their versions. &#x20;
