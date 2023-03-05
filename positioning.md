# XLConnect positioning 

When explaining XLConnect to new prospects, many reactions are ‘ah, is it like X?’. It usually isn’t as XLConnect is a new concept (or new combination of concepts) that fills a new slot between existing tools. 

This answers some frequently asked questions that can help new users position XLConnect within the range of other tools, to help understand what XLConnect is an what it isn't, and how it carves out a new niche that was previously hard to reach. 

## Access/SQL/VBA

One often used way of trying to separate data from logic is to use a relational database like Access of SQL-Server and move the data to and form that through VBA. The first thing to note is that the XLConnect Datalake is a document based data store (which is why we’re not calling it a database as that. 
Easier to model. Relational databases need to thought out very carefully, the number of people in your organisation will typically be very small, the number that can do it well even smaller. 

Complex data like model parameters and pricing parameters are usually very flexible, meaning the database needs to be ‘highly normalised’. It will be quite some work 
The document oriented ‘NoSQL’ way the XLC Datalake stores the data is a better fit with the free format way Excel approaches data. In most cases you only have to select cells in existing models and reports in Excel to carve out the data and logic. This is quite a difference from looping and transforming data to and from relational schema. 

Any data lineage or approval functions need to be built in wich can become quite a drag. It is hard to protect and verify the integrity of a report as its pieces are fragmented across several tables. 
So yes you can do something very similar, but it will be orders of magnitude more work and a much smaller fraction of your existing workforce will ever be able to help. XLC takes a few days training and some practice to get the same benefits. 

## Python

Another technology often used to try and replace Excel are scripting languages, the leading choice being Python. This is definitely a good choice to setup many computational processes, but it does carve out a different niche. 

The number of business people who will be fluent enough to actually contribute to the information and model architecture will be smaller. Since the days of COBOL the dream has been to make everyone a programmer and it has eluded us ever since. My personal attempts to let my team adopt SQL and Python were basically the final motivation to write XLConnect. Even though students at quant studies usually can code to some extent these days, for the existing workforce that number is still much lower , the number who can solve their daily problems with Excel. Maybe there’s a reason for that. 

Text based languages, even one as straight forward as Python, still requires a programmer to structure and solve a problem in their head, work back from the desired solution to the steps required, then type those out in a text based language. The cognitive overhead of doing that is higher than typing and editing a formula in in a cell until it produces the right result, then moving to the next. This will mean that for any workforcem a higher number of people will be able to solve their problems with Excel. 

Also it is not a choice, XLConnect and Python work together very nicely. The Python client library allows you to connect to the Datalake just like the Excel client. Conversely XLConnect can react to events by calling into external functions through function app triggers or plain webhooks. These functions can be implemented in any major language these days. This way you can offload the heavier lifting with loops, scenarios and routing to Python while leaving the deep continuously changing formulas trees and user interfaces to Excel. At any point you can move a step or transformation form one technology to the other and run the process again. This flexibility and potential for rapidly improving processes is very hard to rival. 

For many office pricing, valuation and reporing tasks, it will be less effort and cost to make them in Excel than in Python, and many more (knowledgeable) team members will be able to contribute effectively. That means higher output at lower cost, that’s a better business case. 

## Google Sheets 

XLConnect could be seen as a technology that allows Excel users to do things Google Sheets users take for granted like version control and connecting to external services. Just most Excel users and companies will be quite entrenched in Excel with experience and existing models. Switching will have high costs. Then adding similar features to Excel is a solution. 

## PowerBI/ Tableau/ Qlikview

These tools are amazing in providing insight in (relational) databases filled with large amounts data. Thing is, in a lot of situations you don’t have a database, nor do you have the data yet. These BI tools are great at reporting and analysing data, not good at creating data from scratch. That’s one of the key differences with XLConnect, it can produce data starting from nothing in a very effective and flexible way because it uses Excel. For prototyping and continuous improvement Excel is still without equal. 

Also when formulas other than aggregations need to be applied (cashflow projection from inputs, discounting, solver etc) Excel will be infinitely easier to complete that task with and many more people will be able to do it. For computation heavy work, BI tools will struggle and Excel will shine. 

Having said that, when there is a lot of structured data, these tools are better then Excel. That's why XLConnect data can be pushed into a relational database to allow the combination with other data and reporting with these tools. 

## Cube databases (Vena, Anaplan)

These tools use Excel as a front end for their proprietary database. This means that you can view results form the server, but adding your own formulas in Excel will usually not work very well, calculations should be implemented on the server in some scripting language. This will be very powerful but have a steep learning curve creating dependency on consultants. Also you will not be able to use your current Excel sheets and connect them, in these tool reports will need to be built form scratch and data needs to go through complex conversion to fit into the system.  

XLConnect Datalake is a database designed for Excel. Excel is not the front end to some other technology/ paradigm, the Datalake is a backend to Excel. It fits the free format way Excel deals with data very well so you can connect your existing workbooks without having to change them. 
A second difference is the way data is structured. These cube databases also work very well with financial data that is structured dimensionally (account, period, entity, type, product, customer, etc). This works well for Planning and budgeting results where numbers exist at the crossing of these coordinatates. Data like contracts or model parameters are not dimensioned that regularly 

## Low code frameworks 

These are good at creating data and routing it, but less good at applying formulas to that data. In many applications complex numbers need to be derived in many steps to produce the final outcome. Excel will automatically execute complex tree of formulas on input updates. 

In low code frameworks changes are the execution will need to be scripted step by step. In code with defeats the point of a low code framework. 
You can think of XLConnect as low code framework that can do more, requires less training and can operationalize your existing Excel workbooks. 

## Web applications  

When talking web stack there are a few alternatives that are relatively close to eachother: 

* Asp Core, Angular, Bootstrap, Azure
* Linux, Apache, MySQL, PHP/Perl/Python

The key take away here is that instead of moving the application to a browser, with XLConnect Excel becomes the browser to browse and add to a web based graph of structured data. It’s a bit of a mouth full, let’s unwrap that. 

The table below details the different functions a webstack:

|Function|Web stack|Excel + XLC|
|---|---|---|
|Content|HTML|xlsx|
|Markup|CSS + Bootstrap|xlsx|
|Client side code|javascript|VBA|
|Server side code|C#/ Java/ PHP/ Node.js|Excel Formulas/ VBA / External functions|
|Computations|javascript(?)|Excel Formulas|
|Charts|Charts.js/ Google Charts|Excel charts|
|Client side framework|React/ Vue/ JQuery|XLConnect|
|API's|hand code|XLConnect|
|Hosting|Azure / AWS|XLConnect|
|Database|SQL/ MySQL|XLConnect|
|Authentication|Auth0/ OpenID Connect|XLConnect|
|Transport|Https/ SSL/ Certificates|XLConnect|
|Version control|Git|XLConnect|
|Data Lineage|Program by hand|XLConnect|
|Governance|Program by hand|XLConnect|

The message here is that building a web app takes a lot of knowledge. For every function above there are multiple alternatives where the choice for one will affect the choice for other. Each requires a lot of knowledge and time to master, to be able to build a web app you'll typically a scrum team of 3-4 developers plus a scrum master. Say 5 people, none of them cheap. They will need to learn about the subject of that application (say some financial product) on the job, this takes a lot of communication with room for misunderstanding. The worst thing that can happen is that they almost understand it, that will lead to projects that are 'almost done' for months or however long it takes for someone brave to get a second opinion or just pull the plug. 

Their progress will be measured in sprints of two weeks. For anything trivial, the cost is several (wo)man-months. For anything more complicated start thinking in years. There is a serious cost in terms of time, money and risk associated with building web apps. 

With Excel, a single advanced user will master both Excel and the subject. They'll probably have sample workbooks ready to copy thing over from. They will have working model that produces the correct answer in hours to days. With XLConnect that model can be turned into an application where all the the required technology choices have been made with their best practices integrated into the framework. Instead of mastering a wide variety of technologies (that continuoulsy change) you can now focus on delivering the right answers on time. This means working Line Of Business applications can be deliverd at a fraction of the time, money and risk of doing it through a traditional webapps. 

## Functional Programming

Many of the ideas of XLConnect have been inspired on functional programming concepts. The core idea is that, when inputs and outputs have been defined on a workbook, it becomes a functions from json messages to a jsonmessage. As these messages are free format, this allows for mch more intricate inputs and outputs than people are accustomed to. 

Immutability of versions (of both workbooks and messages) is also taken from functional programming as this unlocks several advantages:

* It ensure consistency of results over time, meaning the audit trail remains valid 
* It makes any result reproducible 
* Approvals work both ways: people can see who approved something and the approver can be sure it cannot be changed afterwards
* It allows for great speedups as they can be cached across the internet and edge devices without the problem of becoming stale 
