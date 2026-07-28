---
description: Pull spend data from the Payhawk API
---

# Payhawk module

Wrappers around the [Payhawk API](https://developers.payhawk.com/) (`auth = 'payhawk'`).

```javascript
payhawk = require('payhawk.js')
payhawk.setAccountId('your-account-id') // required before any call

expenses = payhawk.expenses()
```

## Setup

#### payhawk.setAccountId(id)

Sets the Payhawk account id used to build the API uris. Call this once at the top of your script — the wrappers will throw a 404 without it.

## Wrappers

Each returns the `items` array from the corresponding Payhawk v3 api:

* `payhawk.accounts()`
* `payhawk.accountCodes()`
* `payhawk.cards()`
* `payhawk.expenses()`
* `payhawk.fundAccounts()`
* `payhawk.suppliers()`
* `payhawk.taxRates()`
* `payhawk.teams()`

## Advanced

#### payhawk.callApi(api)

Calls `https://api.payhawk.io/api/v3/accounts/{accountId}/{api}` and returns the `items` array — use this for any endpoint without a wrapper:

```javascript
bills = payhawk.callApi('bills')
```
