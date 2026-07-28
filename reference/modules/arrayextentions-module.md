---
description: 'Experimental: pandas-like statistics and joins for arrays'
---

# ArrayExtentions module

{% hint style="warning" %}
**Experimental** — this module is still in development, has not been tested completely and may change without notice. For the stable everyday methods (sum, avg, agg, unique, ...) use the [SQL module](sql-module.md) instead.
{% endhint %}

Pandas-inspired helpers for arrays and tables, going beyond the SQL module with more statistics and joins.

It can be used two ways:

```javascript
// 1. As methods on every array
ae = require('arrayExtentions.js')
ae.attachExtentions()          // adds the methods to Array.prototype
myTable.median('Amount')

// 2. As plain functions, without touching Array.prototype
ae = require('arrayExtentions.js')
ae.median(myTable, 'Amount')
```

All methods take the same column selectors as the SQL module: a column name string or a lambda.

## Beyond the SQL module

* `median(selector)` / `mode(selector)` / `range(selector)`
* `variance(selector)` / `covariance(sel1, sel2)` / `correlation(sel1, sel2)`

```javascript
r = sales.correlation('AdSpend', 'Revenue')
```

## Reshaping

* `select(columns, aliases)` — project columns, optionally renaming: `rows.select('Name, Age', 'Name, Years')`
* `sortBy(properties)` — sort by one or more columns: `rows.sortBy('Period, AccountCode')`
* `leftJoin(other, leftKey, rightKey)` / `innerJoin(other, leftKey, rightKey)` — join two tables on key columns or lambdas
* Plus the familiar `sum`, `avg`, `min`, `max`, `stdev`, `count`, `unique`, `agg`, `df`, `desc`, `toObject`
