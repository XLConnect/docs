---
description: Pull accounting data from Visma eAccounting
---

# eAccounting module

Wrappers around the [Visma eAccounting API](https://eaccountingapi.vismaonline.com/swagger/ui/index) (`auth = 'eaccounting'`). All wrappers handle paging automatically and return one flat array.

```javascript
ea = require('eaccounting.js')

customers = ea.customers()
balances = ea.accountBalances('2025-09-01')
```

## Wrappers

* `ea.accounts()`
* `ea.accountTypes()`
* `ea.accountBalances(date)` — accepts a date in most common formats, converted to `YYYY-MM-DD` for you
* `ea.articles()`
* `ea.companySettings()`
* `ea.costCenters()`
* `ea.customers()`
* `ea.customerInvoices()`
* `ea.suppliers()`
* `ea.supplierInvoices()`

## Advanced

#### ea.call(api)

Calls any eAccounting api by its relative path, collecting all pages (1000 records per page) into one array. This is how the wrappers above are implemented:

```javascript
vouchers = ea.call('/v2/Vouchers')
```
