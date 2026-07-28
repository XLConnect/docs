---
description: Communicating with the add-in from javascript.
---

# The xlc object

In javascript, you can always access the static xlc object to perform certain actions. This gives you access to low level functionality to enable advanced use cases. Please prefer to use the modules and only drop down to the xlc object when that fails (for instance because you're not working with JSON data, want to make optimized scripts with asynchronous calls, etc).&#x20;

## HTTP Methods

The main task of the xlc object is to enable the javascript engine to perform http request to get and put data to from and to cloud api's.&#x20;

**Please note**: in most cases it is recommended to use a [module](modules/), as that wil wrap a lot of repetitve code. Either the http module or one specific to the cloud system you want to use.&#x20;

### Arguments

* **uri** : string the uri to perform the action on

<pre><code><strong>uri = 'https://api.xero.com/api.xro/2.0/Accounts'
</strong></code></pre>

* **content** : string the content for put and post actions. For example:

```
content = "{id:123,message:'hello'}"

// or rather 
data = {
    id : 123, 
    message : 'hello'
}
content = JSON.stringify(data)
```

* **headers**: object the headers to include in the request, for example:

```
headers = {
    'xero-tenant-id' : '123456asdf'
}
```

* **auth :** string the [cloud system](../developing-with-xlconnect/cloud-systems.md) used, so XLConnet can handle the authentication for you.&#x20;

Auth can be either just the protocol, say 'xero', or it can have a named connection protocol:name. This is useful for certain cloud systems like Visma and Hubspot that put the tenant id in the access token. By using named tokens you can still connect to multiple istances.&#x20;

```
auth = 'xero'

auth = 'xero:clientA'
```

#### get(uri, headers= null, auth = null)

The GET verb gets data from an api.&#x20;

#### put(uri, content, headers= null, auth = null)

#### post(uri, content, headers= null, auth = null)

#### patch(uri, content, headers= null, auth = null)

#### delete(uri, hds, auth)

#### head(uri, hds, auth)

#### options(uri, hds, auth)

### http(method, uri, content, headers, auth)

Http allows for more control than the other methods becasue it also returns the headers where the other mehtods just return the content. Should you have an API interaction that returns usefule information in the headers, use http. It has one extra argument method than can be any of GET, PUT, POST, PATCH, DELETE, HEAD, OPTIONS.

## Scopes&#x20;

#### xlc.requireScope(string auth, string scope)

Logins to most platforms are OAuth variations that have the concept of scopes. The default scopes for each platform are in the settings file. When your workbook needs a special scope you can make sure by using this function.&#x20;

This will check the settings to ensure that scope is in there, if not it will add it to the settings and logout, so the user is sent through the login process again to get a new cookie with these added scope.&#x20;

```
// Make sure the xero login has scope files.read
xlc.requireScope('xero', 'files.read')
```

## Progress Reporting

For longer running scripts, users are a lot happier if they can see things moving and have an ballpark idea of how long they still need to wait. To that end XLConnect shows a progress bar when the script is running, the xlc object has several methods to control what that displays.&#x20;

<figure><img src="../../.gitbook/assets/image (50).png" alt=""><figcaption><p>Progressbar showing the user what is happening and how long it'll take</p></figcaption></figure>

```
xlc.setProgressMessage('Pulling data..')

// do some init stuff
companies = ['A','B','C']
periods = [2022, 2023, 2024]

// set progressbar to total amount of work and set initial message 
xlc.progress(0, 'pulling data..', companies.length * periods.length)

for(company of companies){    
    for(period of periods){
        xlc.progress(-1, `Pulling company ${company} period ${period}..`) // use string interpolation to set progress message 
        // do some api work 
		xlc.sleep(1000)
    }
}
```

#### xlc.progress(int value=-1, string message=null, int max = 0)

This is a newer mehtod that allows for shorthand. As progress reporting can be a significant portion of the code, it helps to be able to do this effectively. the `progress` method allows you to do this will as little code as possible.

```
xlc.progressbar(1, 'Pulling data..', 10) // set max to 10, message and progressbar
xlc.progress() // increments the progressbar by 1 
xlc.progress(10) // set the progressbar to 10
xlc.progressbar(10, 'Pulling Company C..') 


```

#### xlc.setProgressMax(int Max)

This one only sets the Max value for the progressbar.&#x20;

```
xlc.setProgressMax(10)
xlc.setProgressMax(companies.length * periods.length) // assuming companies and periods are 
                                                      // arrays that we will be looping over

```

#### &#x20;setProgressValue(int Value)

```
progress = 0 // declare variable to keep progress 

for(company of companies){

```

setProgressMessage(string Message)

## Datalake&#x20;

These are methods to read and write data to the datalake. These are raw functions that read and write text as a string without interpreting the data. If you are working with json data, consider usign the file module first.&#x20;

#### xlc.fileWrite(string path, string content)

```
dat = { a : 100, b : 200 }  // create little bit of data 
raw = JSON.stringify(dat)   // serialize that data to a string  
path = 'client1/somefolder/dat.json' // relative path under AppData/XLConnect/Data
xlc.fileWrite(path, raw)    // write the data to disk         
```

#### xlc.fileRead(string path)

```
path = 'client1/somefolder/dat.json' // relative path under AppData/XLConnect/Data
raw = xlc.fileRead(path)     // read the raw string from disk 
dat = JSON.parse(raw)        // interpret as json
```

#### xlc.fileDelete(string path)

```
path = 'client1/somefolder/dat.json' // relative path under AppData/XLConnect/Data
xlc.fileDelete(path)     // delete that file
```

#### xlc.fileList(string path, string filter = "\*", bool subfolders = false)

```
folderPath = 'client1/somefolder'
files = xlc.fileList(folderPath) // list all files in this folder 
files = xlc.fileList(folderPath, '*.json') // list all .json files in this folder 
files = xlc.fileList(folderPath, '*.json', true) // list all .json files in this folder and below 
files = xlc.fileList(folderPath, 'null, true) // list all files in this folder and below

data = files.map(xlc.fileRead) // read these files 

```

#### xlc.folderList(string path, string filter = "\*", bool subfolders = false)

#### xlc.folderDelete(string path, bool recursive = false)

## Encoding



#### xlc.SHA256(string content, string privatekey)

#### xlc.base64Encode(string text)

#### xlc.x2j(string xml)

Converts an xml string to equivalent json.

#### xlc.j2x(string json)

Converts&#x20;

## Code Flow&#x20;

#### xlc.sleep(int milliseconds)

Allows you to pause the code. Useful if you have to manually sit out a API rate limit timeout that isn't handled by the HTTP client resilience.&#x20;







##

