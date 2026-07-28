---
description: Daily FX rates for currency conversion in your models
---

# FX module

Serves daily FX rates from XLConnect's public rates database, for currency conversion in multi-currency models.

```javascript
fx = require('fx.js')

rates = fx.get()                         // latest available rates
rate = fx.convert(rates, 'EUR', 'USD')   // EUR -> USD factor
```

#### fx.get(dates)

* `fx.get()` — the latest available rates
* `fx.get(date)` — the rates for a javascript `Date`
* `fx.get([date1, date2, ...])` — an array of rates for an array of `Date`s

Note the arguments must be javascript `Date` objects — use [util.parseAnyDate](util-module.md#util.parseanydate-date) to convert Excel dates or strings first.

#### fx.convert(fx, from, to)

Given one rates object, returns the conversion factor from one currency to another:

```javascript
amountUSD = amountEUR * fx.convert(rates, 'EUR', 'USD')
```

#### fx.getFXRates(toDate, numPeriods)

Rates for a series of month-ends going back from `toDate` — aligned with [xero.periods](xero-module.md#xero.periods-todate-periods) for periodic reporting. Missing dates (weekends, holidays) are filled with the last valid rates.

#### fx.getLastValidRates()

The most recent day with published rates, searching back from today.
