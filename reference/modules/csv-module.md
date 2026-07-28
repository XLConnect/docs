---
description: Convert between CSV text and tables
---

# CSV module

Converts between CSV text and tables (arrays of objects). The module works on strings — combine it with the [FILE](file-module.md), [SPDL](spdl-module.md) or [HTTP](http-module.md) module to actually fetch or store the text.

```javascript
csv = require('csv.js')

raw = xlc.fileRead('imports/prices.csv')  // get CSV text from somewhere
rows = csv.read(raw)                      // -> array of objects
```

#### csv.read(text)

Parses CSV text into an array of objects, using the first line as column names. The delimiter (`,` `;` tab or `|`) and line-break style are detected automatically, and quoted values with embedded delimiters are handled. Empty header cells become `Column0`, `Column1`, ...

#### csv.write(rows, delimiter = ',', headers = null)

Converts an array of objects to CSV text. Columns are the union of all properties found in the rows, unless you pass `headers` (an array) to control the selection and order. Values containing the delimiter or quotes are quoted automatically.

```javascript
text = csv.write(rows, ';')                  // Excel-friendly for EU locales
xlc.fileWrite('exports/prices.csv', text)
```
