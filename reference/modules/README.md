# Modules

Modules are pieces of javascript that are stored outside of the workbook, allowing them to be re-used. This allows code that is used often to be written once and imported into workbooks.&#x20;

You can use the modules that are provided by XLConnect, or you can write and import your own.&#x20;

## Using XLConnect modules

XLConnect comes with a set of modules that encapsulate logic to deal with things like http and the Data Lake, and provide logic that makes it easier to work with cloud platforms like Xero and Visma. Here's how you can add this to your code:

1. Open the Javascript Studio
2.  In the toolbar, drop down the `Include Module` button, then select the `http.js` module.&#x20;

    <figure><img src="../../.gitbook/assets/image (17).png" alt=""><figcaption><p>How to insert the http.js module</p></figcaption></figure>

Afterwards, your Javascript Studio should look like this:

<figure><img src="../../.gitbook/assets/image (12).png" alt=""><figcaption><p>Javascript Studio after inserting the http.js module</p></figcaption></figure>

Notice how a new line of code was added in the Code Pane

```javascript
http = require('http.js')
```

and a new object named `http` with 5 functions appeared on the right.&#x20;

When you use the toolbar to insert a script, it will automatically insert the module definition at the top of the script, execute that line and expand the object that was just created. You can also manually add the line, highlight it, press F8 and expand the new object, we thought we'd save you some clicks as you'll be using this a lot.&#x20;

3. Now drag the `http` function from the Engine Pane onto the code, now your code will look like this:&#x20;

```javascript
http = require('http.js')
http.get(uri, hds = null, auth = null)
```

The http function has three parameters, where the last two have defaults.&#x20;

4. We now need to create a variable `uri`, as the other two parameters have default values (tip you can copy from here if you hover over it, then paste it into the Code Pane):

```javascript
http = require('http.js')
uri = 'https://data.meteoserver.nl/api/uurverwachting_gfs.php?locatie=Amsterdam&key=218a949186'
dat = http.get(uri)
```

5. When we execute this code with F5, it will download the weather forecast for Amsterdam for the next 10 days:

<figure><img src="../../.gitbook/assets/image (18).png" alt=""><figcaption><p>The result of running the code above</p></figcaption></figure>

6. Now double-click the get function to zoom in.&#x20;

Elements in the Engine pane can be zoomed in by double-clicking. When you zoom in on a function, it will show you the code. This can be helpful to understand what parameters it expects from you and what the code does.

<figure><img src="../../.gitbook/assets/image (9).png" alt=""><figcaption><p>Zooming in on the get function to see the code</p></figcaption></figure>

### HTTP Module for making HTTP Calls

The http module can be added with:&#x20;

```
http = require('http.js')
```

### File Module for working with JSON data

```
file = require('file.js') 
```

### Available modules

The full set of modules lives in the public [XLConnect/modules](https://github.com/XLConnect/modules) repository on GitHub:

* **General purpose**: [http](http-module.md), [file](file-module.md), csv, sql, ftp, fx, util, arrayExtentions, spdl
* **Cloud systems**: xero, visma, exact, eaccounting, brincr, hubspot, payhawk

Each cloud system module wraps the http module with authentication and helper methods for that platform.



## Using external modules&#x20;

The javascript world has a rich ecosystem of libraries that you can leverage to manipulate data, here are a few suggestions to get you started:&#x20;

### Lodash

[Lodash ](https://lodash.com/docs)is a versatile JavaScript utility library that offers a rich set of functional programming helpers for handling arrays, objects, and strings without mutating the original data. It simplifies complex operations, enhancing code readability and maintainability.

To import Lodash into your code (please note the version number)

```
_ = require('https://raw.githubusercontent.com/lodash/lodash/4.17.10-npm/lodash.js')
```

### AlaSql

[AlaSQL ](https://github.com/AlaSQL/alasql/wiki)is an open-source SQL database for JavaScript that allows for fast in-memory operations and can work with various data sources, including arrays of objects, Excel, and other databases. It offers a rich SQL-syntax for data manipulation, enabling tasks like querying, joining, and exporting data in a manner familiar to SQL developers.

```
alaSql = require('https://unpkg.com/alasql@4.1.2/dist/alasql.min.js')
```

Where utility libraries like Lodash and Ramda excel at data transformation using chained operations and functional paradigms, AlaSQL shines in querying, aggregating, and joining data sets in a manner reminiscent of relational databases.

### Ramda&#x20;

[Ramda ](https://ramdajs.com/docs/)is a functional programming library for JavaScript that emphasizes pure functions, auto-currying, and immutability, making data transformations concise and predictable. It provides a suite of utilities designed for point-free coding, ensuring functions remain unmodified by side-effects.

```
R = require('https://raw.githubusercontent.com/ramda/ramda/master/dist/ramda.js')
```

In contrast to Lodash, Ramda places a stronger emphasis on immutability and function composition through auto-currying. While both libraries provide utility functions for data manipulation, Lodash leans more towards general utility, whereas Ramda is firmly rooted in functional programming principles.

These are just a few examples to get you started, the general way of importing external modules is through either raw.githubusercontent.com or unpkg.com, that will allow you to grab the source in a single file for the CommonJS module system XLConnect uses today. We are working on supporting ES Modules as well, please let us know if this is important to you.&#x20;



## Developing your own modules



