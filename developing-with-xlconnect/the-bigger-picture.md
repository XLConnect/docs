---
description: Why you'd want to build cloud apps in Excel
---

# The Bigger Picture

## Excel as a client for reading and writing data from cloud APIs

\
XLConnect was designed to build web apps with Excel. If you look at web apps nowadays, they're typically built with a stack of technolgies like HTML, CSS, Javascript (frameworks), server side code, databases, etc. To build a webapp you need to combine quite a few technologies. The HTML/CSS/javascript is served to your browser and it them calls into API's to fetch data.&#x20;

XLConnect comes batteries included with everything you need to read and write data from API's, so now you can build quite similar apps in Excel. The range of required technologies you need to master is a lot smaller, so a single person can do it. The person that understands Excel often also understands the business, removing knowledge transfer to developers. The models you want to automate are probably already made, you just need to wire them up. Lastly Excel is vastly more productive than HTML based tools, especially when it concerns.

<figure><img src="../.gitbook/assets/ExcelvsDev.jfif" alt="" width="375"><figcaption><p>Excel productivity</p></figcaption></figure>

Concluding, building apps, especially for financial modelling and reporting, is vastly more productive in Excel. Having Excel wired up to the web is a big deal.&#x20;

## &#x20;Storing Data in the DataLake

XLConnect also lets you write data to disk, without having to design a database first. Just&#x20;

## How to structure your apps

XLConnect is built as a 'layer cake' of layers to make it all work. The top layer is the Excel workbook, that is designed to be foolproof for users with button. If you go lower you find They use lower layers to make it all work. If you want to, you can go down the layers and replace them with something of your own. The top is intended to be easy to use while giving you the things 80% or users will want. The top layer are build on lower layers, that you can replace with you own code if you want to do so. This requires a bit more technical knowledge but this also gives you full&#x20;

## Deploying Applications to your users



