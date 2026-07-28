---
description: How to sort, filter and join data like you used to with SQL
---

# Working with JSON data

The data you pull from APIs arrives as arrays of objects — think of them as tables, one object per row. Everything you used to do with SQL has a one-line javascript equivalent.

Throughout, assume a table like this:

```javascript
journals = [
    { Account : '0200', Period : '2024-10', Amount :  100 },
    { Account : '0600', Period : '2024-10', Amount : -100 },
    { Account : '0200', Period : '2024-11', Amount :  520 }
]
```

## Filtering (WHERE)

```javascript
oct = journals.filter(j => j.Period === '2024-10')
big = journals.filter(j => j.Amount > 200)
```

## Selecting and computing columns (SELECT)

```javascript
slim = journals.map(j => ({ Account : j.Account, Amount : j.Amount }))
vat  = journals.map(j => ({ ...j, Vat : j.Amount * 0.21 }))   // add a column
```

## Sorting (ORDER BY)

```javascript
sorted = journals.sort((a, b) => a.Account.localeCompare(b.Account))
```

## Aggregating (GROUP BY) — the SQL module

`filter`, `map` and `sort` are built into javascript. For grouping, statistics and lookups, add the [SQL module](../../reference/modules/sql-module.md) — requiring it enables the methods on every array:

```javascript
require('sql.js')

perAccount = journals.agg({
    Account : group,
    Amount  : sum
})

total    = journals.sum('Amount')
accounts = journals.unique('Account')
```

## Joining

The pattern for joining is a lookup object — build it once, then enrich rows as you map over them:

```javascript
require('sql.js')

byCode = accounts.toObject('Code')       // index the accounts table by Code
report = journals.map(j => ({
    ...j,
    AccountName : byCode[j.Account].Name  // JOIN accounts ON Code = Account
}))
```

## From table to sheet and back

Connected Excel tables arrive in your script as exactly this shape — an array of objects, one per table row — and whatever array of objects your script outputs becomes the table's rows. So every technique on this page applies directly to the data your users see.

For a deeper toolkit (median, correlation, joins as one-liners) there is also the experimental [ArrayExtentions module](../../reference/modules/arrayextentions-module.md), and heavy SQL users may enjoy [AlaSQL](../../reference/modules/README.md#alasql) for literal SQL syntax over arrays.
