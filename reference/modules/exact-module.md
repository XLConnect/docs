---
description: Pull financial data from Exact Online
---

# Exact module

Wrappers around the [Exact Online API](https://start.exactonline.nl/docs/HlpRestAPIResources.aspx) (`auth = 'exact'`). The wrappers handle OData paging, strip the metadata from the results, and fill in your current division automatically.

```javascript
exact = require('exact.js')

accounts = exact.GLAccounts()
```

## Division handling

Exact apis are scoped to a division (administration). The wrappers default to the division of the logged-in user; pass a `division` argument to query another one.

#### exact.Me()

Returns the profile of the logged-in user, including `CurrentDivision`.

#### exact.Divisions() / exact.AllDivisions()

Lists the divisions you have access to. Use these to find division codes for multi-administration pulls.

## Wrappers

All take optional `select` (comma-separated fields), `filter` (OData filter) and `division` arguments:

* `exact.FinancialPeriods(select, filter, division)`
* `exact.GLAccounts(select, filter, division)`
* `exact.Journals(select, filter, division)` — defaults to `ReportingYear,ReportingPeriod,GLAccountCode,Amount`
* `exact.ReportingBalance(select, filter, division)` — same default select
* `exact.ProfitLossOverview(select, filter, division)` — returns a single object

```javascript
balance = exact.ReportingBalance(null, "ReportingYear eq 2024")
```

## Other apis

Only the most used apis are wrapped, to keep the module easy to browse. For everything else there are three options with increasing control:

#### exact.callAPi(api, select, filter, division, loginAlias)

Takes the api path following the `{division}/` part of the Resource URI column in the Exact docs. If the Resource URI is `/api/v1/{division}/activities/Complaints`:

```javascript
complaints = exact.callAPi('activities/Complaints')
```

#### exact.query(query)

Paste in a query built with Exact's [REST API Reference tool](https://start.exactonline.nl/docs/HlpRestAPIResources.aspx) — the full path including `/api/v1/...`:

```javascript
res = exact.query('/api/v1/1234567/financial/GLAccounts?$select=Code,Description')
```

#### exact.getAllPages(uri, loginAlias)

Takes a full uri, handles paging and strips the OData metadata. Use this when you need query arguments the wrappers don't build, like `$count` or `$top`.

## Multiple logins

`callAPi` and `getAllPages` take a `loginAlias` argument: pass a [named token](http-module.md#named-auth-parameters) alias to run the call under `exact:<alias>` instead of the default `exact` login.
