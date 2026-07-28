---
description: Using a SharePoint site to store JSON data
---

# Data Lake (SharePoint)

Most companies will have a Office 365 subscription that comes with SharePoint, this is an ideal place to store data that you want to keep private for the company.&#x20;

{% hint style="info" %}
This is the **recommended** way to store data with XLConnect: the syntax is the easiest, and your data stays on your own Office 365 subscription. See the [SPDL module reference](../reference/modules/spdl-module.md) for all available methods.
{% endhint %}

### Quick Demo

Below is a quick demo how to write and read data from a SharePoint drive

<pre><code><strong>spdl = require('spdl.js') // import the spdl module 
</strong>spdl.siteName = 'XLConnectDataLake' // put target SharePoint here (see below)

// create some data 
author = {
	firstName : 'John',
	lastName : 'Doe',
	dob : '2001-01-01',
	books : [
		'The Care And Feeding Of Cats',
		'How To Survive The City'
	]
}

// write it to the data lake 
spdl.write('authors/John-Doe.json', author)

// read it back into a veriable
authorCopy = spdl.read('authors/John-Doe.json')

</code></pre>

As you can see there is very little 'technical friction' to use the SharePoint Data Lake, you don't have to define a database and a table before you can store data and there are no insert/ update/ delete statements, every file is defined by its path and you can just throw it in as is.&#x20;

### Setting up your SharePoint Site

The `XLConnectDataLake` name above is a SharePoint site name. Below is how you can find a SharePoint you have access to.

```
spdl = require('spdl.js')  // import the module 
spdl.sites()  // list available SharePoint sites.
```

This will give you a list of SharePoint sites available

<figure><img src="../.gitbook/assets/image (54).png" alt=""><figcaption></figcaption></figure>

Pick one an assign it to the module:&#x20;

```
spdl.siteName = 'XLConnectDataLake' // replace with your selected target SharePoint site
```

Do this at the top of your code it will remember this for every call

Some SharePoint sites will have multiple drives, in which case it will default to the 'Documents' drive. If you want to be specific you can also point to a specific drive:

### Picking a Drive (Optional)

```
spdl = require('spdl.js')
spdl.siteName = 'XLConnectDataLake'

spdl.drives() // list drives under site
spdl.driveName = 'Documents' // pick drive
```

