---
description: Work with local JSON data on your own computer
---

# FILE module

The file module lets you read and write data to the [Data Lake](../../developing-with-xlconnect/data-lake-xlc.md). Here's a quick example of how to use the write and read methods on the file module:

```
file = require('file.js') // use the file module 

// create an object 
myObject = {
	firstname : 'John',
	lastName : 'Doe',
	dateOfBirth : '1982-06-24'	
}


// save it 
file.write('people/john doe.json', myObject)

// read it back 
obj2 = file.read('people/john doe.json')

```

When you have dealt with SQL databases before, you will find it takes a lot less effort to save data. Especially when dealing with complex data this will become apparent:&#x20;

```
// create array of (simple) journals)
journals = [
	{
		id : 123, 
		date : '2024-10-12',
		lines : [
			{
				account : '0200',
				amount : 100
			},
			{
				account : '0600',
				amount : -100 
			}
		]
	},
	{
		id : 345, 
		date : '2024-10-15',
		lines : [
			{
				account : '0200',
				amount : 520
			},
			{
				account : '0600',
				amount : -520 
			}
		]
	}
]

// write it to disk 
file.write('journals/2024/10.json', myObject)

// read it back
j2 = file.read('journals/2024/10.json')

```

Normally, if you would save this to a sql database, you would first save the first header, capture the generated id, then save the lines with a 'foreign key' pointing to the id of the header. Then save the next line, etc. With the json datalake, you can save the entire set of journals to disk in a single operation, and read them back into memory.&#x20;

