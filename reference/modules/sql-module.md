---
description: Sort, filter, aggregate and join tables like you used to with SQL
---

# SQL module

The sql module adds SQL-like methods to every javascript array, so you can aggregate and reshape the tables you pull from APIs without writing loops. Just requiring it activates the methods — no further setup:

```javascript
require('sql.js')

total = journals.sum('Amount')
```

A *table* here is what the modules return everywhere: an array of objects, one object per row.

## Column selectors

Every method takes a column selector, which can be either a **string** with the column name or a **lambda** computing a value per row:

```javascript
lines.sum('Amount')                        // sum a column
lines.sum(e => e.Amount * e.Quantity)      // sum a computed value
```

## Statistics

* `arr.sum(selector)` — precise even on large arrays (Kahan summation)
* `arr.avg(selector)`
* `arr.min(selector)` / `arr.max(selector)`
* `arr.stdev(selector)` — **sample** standard deviation, matching Excel's `STDEV.S`
* `arr.count(filter)` — number of rows, optionally only those matching a filter lambda

## Reshaping

#### arr.agg(operations)

Group and aggregate in one call — the workhorse. `operations` maps each output column to an operation: `group`, `sum`, `min`, `max`, `count`, `avg` or `stdev`:

```javascript
require('sql.js')

perAccount = journals.agg({
    AccountID : group,
    Period    : group,
    Amount    : sum
})
```

#### arr.unique(selector)

Distinct values of a column (or of the array itself).

#### arr.toObject(keySelector)

Turns a table into a lookup object keyed by a column — fast repeated lookups:

```javascript
byId = accounts.toObject('AccountID')
name = byId['4711'].Name
```

#### arr.df(columns)

Converts a table (array of objects) to a dataframe (object of arrays), optionally only for the named columns: `arr.df('Name, Amount')`.

#### arr.desc()

Describes each column of a table: `{ Column, Min, Max, Sum, Avg, StDev, Count, Unique }` — a quick profile of freshly pulled data.

#### arr.copyColumnsFrom(source, keys, columns)

Carries manually-added columns over from a previous pull onto fresh data, matching rows by key column(s):

```javascript
freshPull.copyColumnsFrom(oldTable, 'Id', ['Notes', 'Category'])
```

Omit `columns` to copy every column that exists in the old table but not the fresh one — i.e. exactly the columns a user added by hand.
