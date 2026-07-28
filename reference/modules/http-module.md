---
description: Make HTTP requests to cloud API's
---

# HTTP module

The HTTP module wraps some logic around making HTTP calls. It usually expects data to be in json form, but can be used with any  return format using the http.http method.&#x20;

For people coming from web development, node.js or deno, this replaces their fetch api as that is a bit too complicated for most users.&#x20;

## How to use in your project&#x20;

If you want to make http calls to web services, consider the http module first. It contains wrapper methods that allow you to focus on the data instead of ceremonial code.&#x20;

On your Javascript Studio window, click the Add Module button and select http

<figure><img src="../../.gitbook/assets/image (2).png" alt=""><figcaption></figcaption></figure>

Add the code below to complete the example

```
http = require('http.js')                // use the http module with wrapper methods 
uri = 'https://swapi.dev/api/people/'    // uri with details about starwars characters
sw = http.get(uri)                       // get the data and assign to sw variable 
```

When you hit F5, this will grab the data from the uri, parse the json that gets returned and give the in-memory javascript object back to you, ready to be used in your code.&#x20;

<figure><img src="../../.gitbook/assets/image (4).png" alt=""><figcaption></figcaption></figure>

## Methods

### Wrapped methods

Please see below the 'wrapped' http methods. They allow you to make http calls with the minimal amount of effort. Please see the paragraphs below for examples how to use the uri, hds, content and auth parameters.&#x20;

* `http.get(uri, hds = null, auth = null)`
* `http.post (uri, content, hds = null, auth = null)`
* `http.put (uri, content, hds = null, auth = null)`
* `http.patch (uri, content, hds = null, auth = null)`
* `http.delete(uri, hds = null, auth = null)`
* `http.head(uri, hds = null, auth = null)`

### HTTP calls with reply headers&#x20;

The calls above handle all (de)serialization automatically, and only return the **response payload**. For most use cases these shorthand operations are the easiest way to handle these requests.&#x20;

Certain API's put data in the reply headers, like the location (uri) of the newly created object. For those scenarios you can use the `http` method:&#x20;

* `http.http(verb, uri, content, hds, auth)`

This does not reply with the deserialized content of the call, but replies with the full http reply including headers and content.&#x20;

```
// make a call
res = http.http('GET', 'https://api.xero.com/api.xro/2.0/accounts', null, { 'xero-tenant-id' : '151ba374-7a79-4691-b1b2-cde45afeb1a1'}, 'xero')

// the content is a string, deserialize to object (if it is JSON). 
dat = JSON.parse(res.Content)

// below the content of http reponse in res:
{
    "StatusCode": 200,
    "Succes": true,
    "Headers": {
    "Xero-Correlation-Id": "f3983711-e8ac-45ea-ad62-c9d20188cc9e",
    "X-AppMinLimit-Remaining": "9997",
    "X-MinLimit-Remaining": "59",
    "X-DayLimit-Remaining": "4900",
    "Pragma": "no-cache",
    "Connection": "keep-alive",
    "X-Client-TLS-ver": "tls1.2",
    "Cache-Control": "no-store, no-cache, max-age=0",
    "Date": "Tue, 27 Aug 2024 11:24:50 GMT",
    "Server": "nginx"
    },
    "Content": "{\r\n  \ ..... "
}

```

## The uri parameter

This is the uri that the call should go to. Typically there will be some base address to the api, and relative addressed for each call.&#x20;

```
baseAddress = 'https://api.xero.com/api.xro/2.0/'
api = 'accounts'
uri = baseAddress + api // build the uri from baseAddress and api 
hds = { 
    'xero-tenant-id' : '151ba374-7a79-4691-b1b2-cde45afeb1a1'
}
auth = 'xero'
accounts = http.get(uri, hds, auth)
```

## The auth parameter

The auth parameter serves two purposes. It determines the authentication to be used, and it allows you to save multiple tokens for that platform. The latter is handy for platforms like Visma.Net that hardcode the system into the token on login, as opposed to systems liek Xero or Exact that have one token and let you decide which organisation to pull in the request.&#x20;

### Default Auth parameter (no name supplied)

The code sample below shows a unnamed auth parameter, it will be stored as 'visma.net' in the manage tokens screen.&#x20;

```
auth = 'visma.net'
http.get(uri, auth)
```

### Switching environments&#x20;

As a consultant you'll often have to switch environments for different clients, here is how to delete a token so the next request for data will take you through the login process again:&#x20;

1.  Go to `Settings > Advanced > Manage Tokens`

    <figure><img src="../../.gitbook/assets/image (52).png" alt="" width="190"><figcaption></figcaption></figure>
2. Right click the cloud system you want to logout from and hit `Delete`.&#x20;

<figure><img src="../../.gitbook/assets/image (53).png" alt="" width="364"><figcaption></figcaption></figure>

When you hit pull (or any operation that contacts the API) it will find no token and send you through the login process again.&#x20;

### Named auth parameters

XLconnect allows you to store multiple tokens of the same cloud platform (protocol) simultaneously, this can save you from have to logout and back in again to switch connections. The sample below uses a 'named token':

```
// the syntax is <protocol>:<token name> 
auth = 'visma.net:client1'
http.get(uri, null, auth)
```

This is handy when having bespoke work for multiple clients, then you can have named tokens to not have to delete the tokens and log in again (and forgetting about that).&#x20;

Or dynamically for integration/ consolidation/ intercompany scenarios:

```
countries = ['NL', 'GB','DE'] // list can also come from Excel worksheet
for(country of countries){
    auth = 'visma.net:' + country 
    dat = http.get(uri, null, auth)
}
```

## The headers parameter

The headers parameter is a javascript object where every property will be a header in the [http request](../../developing-with-xlconnect/javascript/calling-apis-with-http-requests.md#headers).&#x20;

```
uri = 'https://api.xero.com/api.xro/2.0/accounts'
auth = 'xero'
headers = {
    'xero-tenant-id' : 'ad379c37-a74e-4db5-8975-d1d077d03faa' 
}
dat = http.get(uri, headers, auth)
```

## The content parameter

So far we have only shows GET requests that only fetch data from the API, on the differentiating features of XLConnect is that you can also call POST, PUT, PATCH etc http methods, these will typically require a content (or 'payload').&#x20;

