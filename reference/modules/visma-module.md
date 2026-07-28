---
description: Pull business data from the Visma.Net API
---

# Visma module

Wrappers around the [Visma.Net API](https://integration.visma.net/API-index/) (`auth = 'visma.net'`). The wrappers handle paging automatically (1000 records per page) and return one flat array.

```javascript
visma = require('visma.js')

salesOrders = visma.salesOrder()
```

## Wrappers

All take an optional `args` and `entity` argument:

* `visma.account(args, entity)`
* `visma.subAccount(args, entity)`
* `visma.financialPeriod(args, entity)`
* `visma.generalLedgerBalanceV2(args, entity)`
* `visma.salesOrder(args, entity)`
* `visma.purchaseOrder(args, entity)`
* `visma.inventory(args, entity)`
* `visma.journalTransactionV2(args, entity)`

The `args` works like this:

1. **Omitted** — all records are returned: `visma.account()`
2. **A single value** — appended to the url as a path, returning one record: `visma.account(123)`
3. **An object** — appended as query parameters: `visma.account({ active : true })`

The `entity` selects a [named token](http-module.md#named-auth-parameters) — the call runs under `visma.net:<entity>`. Visma.Net hardcodes the company into the access token at login, so named tokens are how you work with multiple companies:

```javascript
countries = ['NL', 'GB', 'DE']
orders = countries.flatMap(c => visma.salesOrder(null, c))
```

## Other apis

#### visma.callAPi(api, args, entity)

Calls any Visma.Net api by its relative path with the same conventions — this is literally how the wrappers are implemented:

```javascript
shipments = visma.callAPi('v1/shipment', { status : 'Open' })
```

#### visma.getAllPages(uri, hds, progressMessage, auth)

Takes a full uri and collects all pages. Use when you build the uri yourself.

#### visma.swaggerSpec()

Returns the swagger spec of the Visma.Net api, so you can browse every api, its parameters and DTO formats from Javascript Studio:

```javascript
spec = visma.swaggerSpec()
spec.paths['/controller/api/v2/salesorder'].post  // browse a specific api
spec.definitions.SalesOrderUpdateDto              // look up a DTO format
```

Note the `DtoValueOfXxx` types in the spec mean the field should be entered as `{ value : 123 }`.

## Service authentication

#### visma.clientCredentials(client\_id, client\_secret, tenant\_id, scope)

Requests a token using the [Client Credentials flow](../../developing-with-xlconnect/cloud-systems.md#visma-client-credentials-service), for background processes that run without a user. `visma.clientCredentialsStaging(...)` does the same against Visma's staging environment.
