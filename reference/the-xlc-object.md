---
description: Communicating with the add-in from javascript.
---

# The xlc object

In javascript, you can always access the static `xlc` object to perform certain actions. This gives you access to low level functionality to enable advanced use cases. Prefer to use the [modules](modules/) and only drop down to the xlc object when they don't cover your case (for instance because you're not working with JSON data, or want to make optimized scripts with asynchronous calls).

**Please note**: the xlc HTTP and message methods return .NET Tasks. Append `.Result` to wait for the call and get the value, like the modules do internally:

```javascript
raw = xlc.get('https://api.xero.com/connections', null, 'xero').Result
dat = JSON.parse(raw) // xlc returns raw strings, parsing is up to you
```

## Properties

#### xlc.version

The version of the installed add-in, as a string.

#### xlc.personalDatabase

The name of your personal database.

## HTTP Methods

The main task of the xlc object is to enable the javascript engine to perform http requests to get and put data from and to cloud APIs.

**Please note**: in most cases it is recommended to use a [module](modules/), as that will wrap a lot of repetitive code. Either the http module or one specific to the cloud system you want to use.

### Arguments

* **uri** : string the uri to perform the action on

<pre><code><strong>uri = 'https://api.xero.com/api.xro/2.0/Accounts'
</strong></code></pre>

* **content** : string the content for put and post actions. Note this is a string, so serialize your data first:

```javascript
data = {
    id : 123, 
    message : 'hello'
}
content = JSON.stringify(data)
```

* **headers**: object the headers to include in the request, for example:

```javascript
headers = {
    'xero-tenant-id' : '123456asdf'
}
```

* **auth :** string the [cloud system](../developing-with-xlconnect/cloud-systems.md) used, so XLConnect can handle the authentication for you.

Auth can be either just the protocol, say 'xero', or it can have a named connection protocol:name. This is useful for certain cloud systems like Visma and Hubspot that put the tenant id in the access token. By using named tokens you can still connect to multiple instances.

```javascript
auth = 'xero'

auth = 'xero:clientA'
```

### Methods

All methods return the response content as a string (remember `.Result`), except `http` which returns the full reply.

#### get(uri, headers = null, auth = null)

Gets data from an api.

#### put(uri, content, headers = null, auth = null)

Sends `content` with the PUT verb, returns the reply content.

#### post(uri, content, headers = null, auth = null)

Sends `content` with the POST verb, returns the reply content.

#### patch(uri, content, headers = null, auth = null)

Sends `content` with the PATCH verb, returns the reply content.

#### delete(uri, headers = null, auth = null)

Performs a DELETE request, returns the reply content.

#### head(uri, headers = null, auth = null)

#### options(uri, headers = null, auth = null)

### http(method, uri, content, headers, auth)

Http allows for more control than the other methods because it also returns the status code and headers where the other methods just return the content. Should you have an API interaction that returns useful information in the headers, use http. It has one extra argument method that can be any of GET, PUT, POST, PATCH, DELETE, HEAD, OPTIONS.

The reply is a string containing a JSON document with `StatusCode`, `Succes`, `Headers` and `Content` properties. (Note the spelling of `Succes` — this is the actual field name in the product.)

```javascript
raw = xlc.http('GET', 'https://api.xero.com/connections', null, null, 'xero').Result
res = JSON.parse(raw)
if (res.Succes) {
    dat = JSON.parse(res.Content) // Content is a string, parse if it is JSON
}
```

### Rate limits

When an API replies with HTTP 429 (rate limit hit), XLConnect reads the reply's retry timeout, waits it out while showing the countdown to the user, and retries the call once automatically. This works out of the box for Xero and Exact.

## Scopes

#### xlc.requireScope(string auth, string scope)

Logins to most platforms are OAuth variations that have the concept of scopes. The default scopes for each platform are in the settings file. When your workbook needs a special scope you can make sure it is granted by using this function.

This will check the settings to ensure that scope is in there, if not it will add it to the settings and logout, so the user is sent through the login process again to get a new token with the added scope.

Supported auth values: `xero`, `visma.net` and `microsoft`.

```javascript
// Make sure the xero login has scope files.read
xlc.requireScope('xero', 'files.read')
```

## Progress Reporting

For longer running scripts, users are a lot happier if they can see things moving and have a ballpark idea of how long they still need to wait. To that end XLConnect shows a progress bar when the script is running, the xlc object has several methods to control what that displays.

<figure><img src="../.gitbook/assets/image (50).png" alt=""><figcaption><p>Progressbar showing the user what is happening and how long it'll take</p></figcaption></figure>

```javascript
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

This is a newer method that allows for shorthand. As progress reporting can be a significant portion of the code, it helps to be able to do this effectively. The `progress` method allows you to do this with as little code as possible. Leave `value` at -1 to increment the bar by one, pass `message` to update the text and `max` to set the total amount of work.

```javascript
xlc.progress(1, 'Pulling data..', 10) // set progressbar to 1 out of 10 and set the message
xlc.progress() // increments the progressbar by 1 
xlc.progress(10) // set the progressbar to 10
xlc.progress(-1, 'Pulling Company C..') // increment by 1 and update the message
```

#### xlc.setProgressMax(int Max)

This one only sets the Max value for the progressbar.

```javascript
xlc.setProgressMax(10)
xlc.setProgressMax(companies.length * periods.length) // assuming companies and periods are 
                                                      // arrays that we will be looping over
```

#### xlc.setProgressValue(int Value)

Sets the progressbar to a specific value.

```javascript
progress = 0 // declare variable to keep progress 

for(company of companies){
    // do some api work
    progress++
    xlc.setProgressValue(progress)
}
```

#### xlc.setProgressMessage(string Message)

Sets the message shown above the progressbar.

```javascript
xlc.setProgressMessage('Pulling data..')
```

## Datalake

These are methods to read and write data to the datalake. These are raw functions that read and write text as a string without interpreting the data. If you are working with json data, consider using the [file module](modules/file-module.md) first.

All paths are relative to the XLConnect data folder (`AppData/XLConnect/Data`).

#### xlc.fileWrite(string path, string content)

```javascript
dat = { a : 100, b : 200 }  // create little bit of data 
raw = JSON.stringify(dat)   // serialize that data to a string  
path = 'client1/somefolder/dat.json' // relative path under AppData/XLConnect/Data
xlc.fileWrite(path, raw)    // write the data to disk         
```

#### xlc.fileRead(string path)

```javascript
path = 'client1/somefolder/dat.json' // relative path under AppData/XLConnect/Data
raw = xlc.fileRead(path)     // read the raw string from disk 
dat = JSON.parse(raw)        // interpret as json
```

Passing `null` as the path opens a file picker dialog so the user can select the file to read.

#### xlc.fileDelete(string path)

```javascript
path = 'client1/somefolder/dat.json' // relative path under AppData/XLConnect/Data
xlc.fileDelete(path)     // delete that file
```

#### xlc.fileList(string path, string filter = "\*", bool subfolders = false)

```javascript
folderPath = 'client1/somefolder'
files = xlc.fileList(folderPath) // list all files in this folder 
files = xlc.fileList(folderPath, '*.json') // list all .json files in this folder 
files = xlc.fileList(folderPath, '*.json', true) // list all .json files in this folder and below 
files = xlc.fileList(folderPath, null, true) // list all files in this folder and below

data = files.map(xlc.fileRead) // read these files 
```

#### xlc.folderList(string path, string filter = "\*", bool subfolders = false)

Lists the folders under `path`, same arguments as `fileList`.

#### xlc.folderDelete(string path, bool recursive = false)

Deletes the folder at `path`. When `recursive` is false the folder must be empty; pass true to also delete its contents.

## Binary files

#### xlc.downloadBinary(string uri, string path, string auth = null)

Downloads a file from the internet and saves it under the data folder at `path`. Returns the full local path of the saved file.

```javascript
localPath = xlc.downloadBinary('https://example.com/report.pdf', 'downloads/report.pdf')
```

#### xlc.uploadBinary(string uri, string localPath, dynamic headers = null, string auth = null)

Reads the file at `localPath` (a full path, this may be outside the data folder) and PUTs the bytes to `uri`. Returns the reply content.

#### xlc.postVismaAttachment(string uri, string filePath, string auth)

Dedicated method to post a file (relative to the data folder) as a multipart attachment to Visma.net.

## SFTP

Methods to exchange files with SFTP servers. Connect once, then use the other methods.

#### xlc.sftpConnect(string server, int port, string username, string password)

#### xlc.sftpDisconnect()

#### xlc.sftpList(string path)

#### xlc.sftpGet(string path)

Reads the remote file and returns its content as a string.

#### xlc.sftpPut(string remotePath, string content)

Writes a string to a remote file.

#### xlc.sftpDownload(string remotePath, string localPath)

#### xlc.sftpUpload(string localPath, string remotePath)

#### xlc.sftpDelete(string path)

```javascript
xlc.sftpConnect('sftp.example.com', 22, 'user', 'secret')
files = xlc.sftpList('/outbound')
raw = xlc.sftpGet('/outbound/prices.json')
xlc.sftpDisconnect()
```

## Encoding & Hashing

#### xlc.SHA256(string content, string privatekey)

Computes an HMAC-SHA256 signature of `content` with `privatekey`, returned as a base64 string. Some APIs require requests to be signed this way.

Note this is different from `xlc.sha256(text)` below, which is a plain hash.

#### xlc.md5(string text)

#### xlc.sha1(string text)

#### xlc.sha256(string text)

#### xlc.sha512(string text)

Compute the respective hash of `text`, returned as a lowercase hex string.

#### xlc.base64Encode(string text)

Base64 encodes a string.

#### xlc.x2j(string xml)

Converts an xml string to equivalent json.

#### xlc.j2x(string json)

Converts a json string to equivalent xml.

#### xlc.guid()

Returns a new GUID as a string.

## Code Flow

#### xlc.sleep(int milliseconds)

Allows you to pause the code. Useful if you have to manually sit out an API rate limit timeout that isn't handled by the HTTP client resilience.

#### xlc.heapInfo()

Returns the javascript engine's runtime heap statistics, for diagnosing memory usage of heavy scripts.

## Messages

Methods to read and write messages in the XLConnect database directly from code.

**Please note**: for new work we recommend storing data in the [SharePoint Data Lake](../developing-with-xlconnect/data-lake-sharepoint.md) instead — the `spdl` module has an easier syntax and keeps your data on your own Office 365 subscription.

A message is a JSON document with an envelope. The fields you work with from code:

* **db** : string the database the message lives in (see `xlc.personalDatabase` for your own)
* **par** : string the parent folder location of the message
* **typ** : string the message type, for example `Sales.Order`
* **key** : the key that identifies the message within its type, can be a string, number or object
* **dat** : object the actual data payload
* **\_id**, **rev**, **cre**, **crb** : set by the server — id, revision, created timestamp and creator

#### xlc.msg(string db, string id)

Gets a single message by its `_id`, as a json string. Remember `.Result`:

```javascript
// list the folder to find the message, then fetch it by _id
raw = xlc.msgs(xlc.personalDatabase, 'sales/orders').Result
orders = JSON.parse(raw)

raw = xlc.msg(xlc.personalDatabase, orders[0]._id).Result
order = JSON.parse(raw)
console.log(order.dat) // the payload
```

#### xlc.msgs(string db, string loc)

Gets all messages in the folder `loc`, as a json string containing an array.

```javascript
raw = xlc.msgs(xlc.personalDatabase, 'sales/orders').Result
orders = JSON.parse(raw)
```

#### xlc.msgs(string db, string loc, keys)

Gets the messages in `loc` with the given keys, an array. Use this to fetch a specific selection in one call:

```javascript
keys = ['2024-001', '2024-002', '2024-003']
raw = xlc.msgs(xlc.personalDatabase, 'sales/orders', keys).Result
orders = JSON.parse(raw)
```

#### xlc.postMessages(msgs)

Posts an array of message objects to the database in one call. Every message must have `db`, `par`, `key`, `typ` and `dat` properties:

```javascript
msgs = [
    {
        db  : xlc.personalDatabase,
        par : 'sales/orders',
        typ : 'Sales.Order',
        key : '2024-004',
        dat : { customer : 'Acme', amount : 1250.00 }
    },
    {
        db  : xlc.personalDatabase,
        par : 'sales/orders',
        typ : 'Sales.Order',
        key : '2024-005',
        dat : { customer : 'Globex', amount : 730.50 }
    }
]

xlc.postMessages(msgs)
```
