---
description: JavaScript Object Notation is the lingua franca of the API web
---

# JSON Data Format

JSON (JavaScript Object Notation) is how virtually every modern API writes down data — and it's what travels between your workbook, the cloud and the data lake. The good news: it's also literal javascript, so in your scripts there is no translation step. Five minutes here covers all of it.

## The building blocks

**Values** are strings, numbers, booleans or `null`:

```json
"Acme Ltd"     42     19.95     true     null
```

**Objects** hold named values in `{ }` — like one row, or one form:

```json
{
    "Account" : "0200",
    "Name"    : "Sales",
    "Active"  : true
}
```

**Arrays** hold ordered lists in `[ ]`. An array of objects is a table:

```json
[
    { "Account" : "0200", "Amount" :  100 },
    { "Account" : "0600", "Amount" : -100 }
]
```

**Nesting** is where JSON beats flat tables: a value can itself be an object or array. An invoice carries its own lines — no foreign keys, no second table:

```json
{
    "id"    : 123,
    "date"  : "2024-10-12",
    "lines" : [
        { "account" : "0200", "amount" :  100 },
        { "account" : "0600", "amount" : -100 }
    ]
}
```

## JSON in your scripts

In a script, JSON text and javascript data convert with two built-ins:

```javascript
obj = JSON.parse(text)        // text -> data you can work with
text = JSON.stringify(obj)    // data -> text to store or send

obj.lines[0].amount           // drill in with dots and brackets
```

You rarely call these yourself: the [modules](modules/README.md) parse API replies and serialize your data automatically. You only reach for `JSON.parse`/`stringify` with the raw [xlc methods](the-xlc-object.md), which deal in plain strings.

## Dates

JSON has no date type — dates travel as strings, Excel uses serial numbers, and javascript has `Date` objects. The [UTIL module](modules/util-module.md) converts between all three.
