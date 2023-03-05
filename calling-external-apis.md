# Calling External API&#39;s

XLConnect can call external REST API&#39;s in a way that is much more flexible than using Excel&#39;s built in Webservice() function, as XLConnect&#39;s functionality allows for headers, all HTTP methods like Post, Put and uses the cell connections to put the reply into cells on the worksheet (on top of standard XLC functionalities like workbook version management and the ability to create workflows with the output of the webservices.

To add a webservice, click on the Connect API button to bring up it&#39;s configuration screen:

## Anatomy of an HTTP Request

Before we dive into the specific parts of the XLConnect features, it is good to have a quick recap of how calls to an HTTP server work. An HTTP call has two basic parts, a Request and a Reply, that both have quite a few ways to put data in them. Thiese are the parts of an HTTP call.

- Request: a client sends a request to a server that consists of
  - Url
    - Address: the address of the api, for instance
    - Parameters
  - Request Headers: multiple name value-pairs
  - Method: Get or Post
  - Content Headers: multiple name value-pairs
  - Content (only when method is Post)
- Reply
  - Response Code
  - Headers
  - Content

As you can see there are quite a few ways to make an HTTP call to an API, and different API&#39;s use different features. XLConnect can configure every aspect of an HTTP call to make sure we can connect to any API.

Let&#39;s start with a simple call.

## Basic Get Call

The simplest rest call only has an url (usually with parameters) that returns a json document. We will use this one:

[https://data.meteoserver.nl/api/uurverwachting\_gfs.php?locatie=Utrecht&amp;key=218a949186](https://data.meteoserver.nl/api/uurverwachting_gfs.php?locatie=Utrecht&amp;key=218a949186)

if you click the link above you will see the sample json reply that we will work with in this tutorial.

To call this API from XLConnect and display the results, do the following steps:

- Open a new workbook
- Click the button Add Api. This will bring up the Maps &amp; Messages pane with a new API

- Paste the url above into the box url.
- Click the button &#39;Reply&#39; to call the API to get a reply and create (yet unmapped) connections for the fields it returns.

Your screen should now look like this:

![](RackMultipart20200606-4-10x3hpl_html_588de27a5e04b33f.png)

- Drag the following fields onto the indicated cells on the worksheets to connect them.
  - Plaats: B3
  - tijd\_nl: B5
  - temp: C5
  - winds: D5
  - windrltr: E5

Your screen should look like this:

![](RackMultipart20200606-4-10x3hpl_html_594c76d9db6b56a8.png)

- Double-click the Request Map to call the API.

Your screen should now look like this:

![](RackMultipart20200606-4-10x3hpl_html_48595d6eac02e2e6.png)

The json that was requested from the API has been downloaded, and selected values are entered into the designated cells.

Congratulations, you just made you first API call from XLConnect!

## Variable Parameters with Placeholders

If you look at the parameter part of the url, you will see the city &#39;Utrecht&#39; is hardcoded in there. We will now make this a variable so we can check the forecast for other cities.

- Keeping the workbook from the previous exercise, change the url into this:

[https://data.meteoserver.nl/api/uurverwachting\_gfs.php?locatie= **{City}** &amp;key=218a949186](https://data.meteoserver.nl/api/uurverwachting_gfs.php?locatie=%7BCity%7D&amp;key=218a949186)

Please note that the word **Utrecht** in the original has been changed into **{City}**. Note how this ads a new map &#39;UrlAndHeader&#39; below the API with a connection &#39;City&#39;

Your screen should look like this:

![](RackMultipart20200606-4-10x3hpl_html_751fd93e8fd685c0.png)

- Drag the connection &#39;City&#39; onto cell C1 to connect it to that cell
- Type the value &#39;Maastricht&#39; in cell C1
- Double-click the API to call it. You will now see the weather forecast for Maastricht!
- Type different Dutch cities into cell C1 to get localized weather forecast.

## Using Live Keys

The live keys feature, that reacts when cells in the workbook have changed, also works for APIs. This way, when you type a new value into the cell &#39;City&#39;, the forecast will update automatically.

- Richt-click the &#39;City&#39; property under the Request Map and set it as key. The icon for the City connection should ow be a key.
 ![](RackMultipart20200606-4-10x3hpl_html_882297bfd7a4b582.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif) ![](RackMultipart20200606-4-10x3hpl_html_60ddf2a9bb1ee954.gif) ![](RackMultipart20200606-4-10x3hpl_html_199f2d8a861f6de6.gif)

![](RackMultipart20200606-4-10x3hpl_html_f2db5af5f6e0e7fe.png) ![](RackMultipart20200606-4-10x3hpl_html_69e57da5c84ef406.png)

- Click on the Request Map and properties to check that Live Keys is enabled.
 ![](RackMultipart20200606-4-10x3hpl_html_1a1bc788f3747941.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif) ![](RackMultipart20200606-4-10x3hpl_html_f061b0de51fb0b23.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif) ![](RackMultipart20200606-4-10x3hpl_html_3fce91bfebbc350e.gif) ![](RackMultipart20200606-4-10x3hpl_html_833f04d9f99d8411.gif)



![](RackMultipart20200606-4-10x3hpl_html_937330511c1aee04.png)

- Change the value of City (cell C1) into Rotterdam

The API should update automatically when you change the city.

## Headers

Some API&#39;s require values to be sent in the headers, most often for authentication purposes and specififying output format. As you will see standard header &#39;Accept:application/json&#39; is added. This tells the server that we want to receive the output as json. Not every API actually requires or uses this header, but it helps novice users get a successful call when they do.

For this example we will use another API:

[https://rapidapi.com/coingecko/api/coingecko](https://rapidapi.com/coingecko/api/coingecko)

This API returns data about cryptocurrencies. It is hosted on the RapidApi platform that uses headers to authenticate requests. To be able to use the API, you will need to create a free account.

- Go to rapidapi.com and click the &#39;Sign up&#39; button in the top right corner to create your account.
- Then go to this api: [https://rapidapi.com/coingecko/api/coingecko](https://rapidapi.com/coingecko/api/coingecko)
-

## Making a Post call

Besides asking for data, HTTP calls can also send data. This mean XLConnect can make complicated request and write data to specific target systems.

There are two different ways of making HTTP Post calls:

- Generic: call REST API&#39;s that were designed without XLConnect in mind. You construct a JSON document using connections to workbook cells, double click the API connection to extract these values into a JSON document, post that to the API and receive back a JSON document that you can map to worksheet cells and store in the DataLake as a XLConnect Messagem using the JSON document as the body
- XLConnect: call API&#39;s that were designed to be called using an XLConnect message. Instead of typing values in cells, you can double-click a message, this message is then sent to the API for processing, optionally receiving back a XLConnect Message in return.

## Generic API&#39;s

For this exercise, we will use the &#39;Mirror&#39; api that XLConnect provides on this address:

[https://xlconnectfunctionappservice.azurewebsites.net/api/Mirror](https://xlconnectfunctionappservice.azurewebsites.net/api/Mirror=)

This API returns whatever message it received.
