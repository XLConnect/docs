---
description: Date helpers for the three date worlds you juggle in scripts
---

# UTIL module

Scripts constantly juggle three kinds of dates: javascript `Date` objects, Excel date numbers (like `45132`) and ISO strings (`'2023-07-25'`). The util module converts between them.

```javascript
util = require('util.js')

util.YYYY_MM_DD(45132)   // '2023-07-25' — straight from an Excel cell
```

## The universal parser

#### util.parseAnyDate(date)

Takes a string, an Excel date number or a `Date` and returns a `Date`. Most other functions accept any of the three via this parser — pass whatever you have.

## Formatting

* `util.YYYY_MM_DD(anyDate)` — `'2023-07-25'`, the format most APIs want
* `util.YYYY_MM(anyDate)` — `'2023-07'`, handy as a period key
* `util.isoDate(year, month, day)` — build an ISO string from parts

## Converting

* `util.parseExcelDate(number)` / `util.toExcelDate(jsDate)` — Excel serial numbers ↔ `Date`
* `util.parseUnixDate(number)` / `util.toUnixDate(jsDate)` — Unix timestamps ↔ `Date`

## Date math

#### util.eOMonth(startDate, months)

The last day of the month, `months` after `startDate` — the equivalent of Excel's `EOMONTH`, accepting any date format:

```javascript
periodEnd = util.eOMonth('2024-01-15', 0)   // 2024-01-31
nextClose = util.eOMonth(45132, 1)          // end of next month
```

#### util.utcCToday()

Today at midnight UTC, for stable date comparisons.
