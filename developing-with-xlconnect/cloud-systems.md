---
description: List of cloud systems
---

# Cloud Systems

XLConnect comes 'batteries included' with authentication to cloud systems built in, so you can get straight to the business logic.&#x20;

When making [HTTP ](../reference/modules/http-module.md)calls, the final parameter is `auth` to indicate the cloud system you want to authenticate to.&#x20;

Below is a simple example of a direct http call, please note that these cloud systems also have a [javascript module](../reference/modules/) with helper methods.&#x20;

```
http = require('http.js')

// grab data from the xero api
uri = 'https://api.xero.com/connections'
auth = 'xero'
http.get(uri, null, auth)

// or shorthand 
http.get('https://api.xero.com/connections', null, 'xero')
```

## Xero

`auth = 'xero'`

## Visma.Net Native

`auth = 'visma.net'`

Use this authentication for the [Visma.Net API](https://integration.visma.net/API-index/) and the [Visma.net.ERP.SalesOrders.Api](https://salesorder.visma.net/swagger/index.html).&#x20;

This uses the Authentication Code Interactive flow that automatically open a web browser to log the user in and return to the app.&#x20;

## Visma Client Credentials (Service)

`auth = 'vismacc'`  &#x20;

Visma Client Credentials authentication run without user login flow, best fit for background processes that run without a user.&#x20;

A fixed ClientId and Secret are entered once and stored securely, then the code can run without a user having to be present to login.&#x20;

Services Visma API's like [Brincr ](https://sandbox-api.brincr.com/api/documentation)use a Service (Client Credentials flow) authentication.&#x20;

Please note these need to be set up per project in the [Visma Developer Portal](https://oauth.developers.visma.com/).&#x20;

## Visma Client Credentials Multi

This is a second variant of Visma Client Credentials (Service) application Type, that allows the code to switch between tenants under an app registration (the one above is fixed to a single tenant at login).&#x20;

1. Set up a Visma Client Credentials app in the [Visma Developer Portal](https://oauth.developers.visma.com/).&#x20;
2. Publish it to the App Store as 'Invitation Only'&#x20;
3. Generate an Invitation Code (in the developer portal)
4. Go to the [Visma App Store](https://apps.visma.com/), log in, and go to My Apps
5. Select your organisation in the top right corner and click 'Use Invitation Code'
6. &#x20;Repeat steps 3 and 5 to add all required tenants are added under the Application.&#x20;

To connect to a certain tenant:&#x20;

```
// auth for vismacc has three arguments: protocol, clientId and tenantId

clientId = 'isv_***' // enter the Client ID for your application from the Developer Portal
tenantId = 'a0b1100c-4a02-11ed-98d8-06a0b4eada6b' // can be found in Developer Portal > Integration > Customers

auth = `vismacc.multi:${clientId}:${tenantId}`

//for example 
auth = 'vismacc.multi:isv_my_private_app:a0b1100c-4a02-11ed-98d8-06a0b4eada6b'
```

When pulling data for the first time, XLConnect will ask for the Client Secret and Scopes that will be stored securely. After that the code can switch between tenants to connect to all the different tenants under that Application Registration.&#x20;

## Visma eAccounting

```
http = require('http.js')

base = 'https://eaccountingapi.vismaonline.com/'
uri = base + 'v2/accountbalances/2025-09-01'

auth = 'eaccounting'

http.get(uri, null, auth)
```

API description can be found here: \
[https://eaccountingapi.vismaonline.com/swagger/ui/index#](https://eaccountingapi.vismaonline.com/swagger/ui/index)

## Exact Online

`auth = 'exact'`

[`https://start.exactonline.nl/docs/HlpRestAPIResources.aspx?SourceAction=10`](https://start.exactonline.nl/docs/HlpRestAPIResources.aspx?SourceAction=10)

## QuickBooks Online

`auth = 'qbo'`

API description can be found here: [https://developer.intuit.com/app/developer/qbo/docs/api/accounting/all-entities/account](https://developer.intuit.com/app/developer/qbo/docs/api/accounting/all-entities/account)

## Payhawk

`auth = 'payhawk'`

API description here: [https://developers.payhawk.com/](https://developers.payhawk.com/)

## Hubspot

`auth = 'hubspot'`

[https://developers.hubspot.com/docs/api/crm/understanding-the-crm](https://developers.hubspot.com/docs/api/crm/understanding-the-crm)

## Microsoft Graph

Microsoft Graph lets you connect to basically all your data on microsoft including email, Sharepoint, OneDrive etc.&#x20;

`auth = 'ms.graph'`

Documentation can be found here: [https://developer.microsoft.com/en-us/graph/graph-explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)

## Microsoft Dynamics 365

`auth = 'ms.dynamics'`

## Basic Authentication

`auth = 'basic'`

This is the classic Basic Authentication that is still used for older systems.&#x20;



## NMBRS

`auth = 'nmbrs'`



