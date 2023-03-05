# Azure Functions

Azure is Microsoft&#39;s cloud platform, it has many, many different offerings, one of which is Azure Functions. These allow developers to write functions that can do heavy lifting (financial simulations, GPU, machine learning, data science, workflows) without having to worry about the underlying hardware, they&#39;re so called &#39;serverless functions&#39;.

Azure Functions are quite easy to write and deploy, but then actually calling them securely is more work than writing them, because you need to write either a web or desktop app to do so (or call them from a script which is hard to deploy to users), with a database behind it, and deal with complicated web security protocols.

XLConnect and Azure Functions complement each other very well. Excel+ XLConnect provide a very intuitive and productive platform to write and deploy user interfaces, define and gather data, then call the Azure Functions securely to do any heavy lifting.

Because XLConnect is built around the latest web standards for Authentication and Authorization, this can be integrated into Azure Functions with a few lines of code, as we will demonstrate. This relieves developers from a lot of complicated and risky technical details, freeing up time and attention to focus on business requirements.

Because these Azure Functions can also call back into the XLConnect platform, they can be used to create something akin to stored procedures; complex encapsulated logic that the user can call from Excel as a single command. Connecting to Azure Functions opens up an effectively limitless array of capabilities and processing power.

Please note that while this chapter focusses on Azure Functions, the same techniques can be applied for AppServices or Virtual Machines pretty much unchanged, Azure Functions are more likely to appeal to Excel developers because one does not have to deal with many of the technical barriers. This chapter will show examples in C#, but the same programs can also be written in Python of JavaScript (node.js) which are likely to be familiar with business minded developers.
