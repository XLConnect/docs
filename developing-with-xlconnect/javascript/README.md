---
description: Script your workbooks in the language of the data web
---

# Javascript

Scripts in XLConnect are written in javascript, running on the same V8 engine that powers Chrome — inside Excel. If you can write an Excel formula, the step to a few lines of javascript is smaller than you'd think, and it unlocks everything the platform can do: calling APIs, reshaping data, reading and writing the data lake.

A taste — this is a complete, working data pull:

```javascript
http = require('http.js')
accounts = http.get('https://api.xero.com/api.xro/2.0/Accounts', hds, 'xero')
```

This section takes you from zero to productive:

* [Javascript Studio](javascript-studio.md) — the editor built into Excel: run code with `F5`, inspect every variable, zoom into your data.
* [Javascript language](javascript-language.md) — why javascript is the right tool for API data.
* [Calling API's with HTTP Requests](calling-apis-with-http-requests.md) — how requests work and how XLConnect wraps them.
* [Working with JSON data](working-with-json-data.md) — sort, filter, aggregate and join, like you used to with SQL.
* [Advanced Javascript](advanced-javascript.md) — for when the basics aren't enough.
* [Standalone mode](standalone-mode.md) — run scripts outside Excel, on a schedule.

For looking things up while coding, keep the [Reference](../../reference/the-xlc-object.md) close: the xlc object and the [modules](../../reference/modules/README.md).
