---
description: Read and write JSON data on your SharePoint site — the recommended data lake
---

# SPDL module (SharePoint Data Lake)

The spdl module reads and writes files on a SharePoint site through the Microsoft Graph API, turning your Office 365 subscription into a [data lake](../../developing-with-xlconnect/data-lake-sharepoint.md). It handles authentication (via the `spdl` cloud system), OData paging, and serialization for you.

```javascript
spdl = require('spdl.js')
spdl.siteName = 'XLConnectDataLake' // your SharePoint site, see below

spdl.write('authors/John-Doe.json', { firstName : 'John', lastName : 'Doe' })
author = spdl.read('authors/John-Doe.json')
```

## Setup

#### spdl.siteName

Assign the name of your SharePoint site once, at the top of your script. This looks up the site, remembers it for every call, and selects its Document library drive automatically.

```javascript
spdl.siteName = 'XLConnectDataLake'
```

If no Document library can be found automatically, assign the drive yourself with `spdl.driveName` or `spdl.driveId`.

#### spdl.sites()

Lists the SharePoint sites you have access to, with `name`, `displayName`, `description` and `webUrl`. Use this to find the value for `siteName`.

#### spdl.drives()

Lists the drives under the current site. Use this when a site has multiple document libraries and you need to pick one:

```javascript
spdl.siteName = 'XLConnectDataLake'
spdl.drives()                 // list drives under the site
spdl.driveName = 'Documents'  // pick one by name
```

## Reading and writing

Paths are relative to the root of the selected drive. Where a `path` argument is accepted, you can pass either a string or an object with a `path` property — so rows from `spdl.list()` can be passed straight in.

#### spdl.read(path)

Reads a file and returns the parsed data.

#### spdl.write(path, data)

Serializes `data` and writes it to `path`. Folders in the path are created as needed.

#### spdl.readP(paths)

Reads an array of paths **in parallel** — much faster than a loop of `read` calls when loading many files:

```javascript
files = spdl.list('journals/2024')          // list the folder
datas = spdl.readP(files)                   // read all files in parallel
```

#### spdl.writeP(paths, datas)

Writes multiple files in parallel. `paths` is an array of strings, `datas` an array of values of the same length.

```javascript
paths = ['journals/2024/10.json', 'journals/2024/11.json']
datas = [octoberJournals, novemberJournals]
spdl.writeP(paths, datas)
```

## Browsing

#### spdl.list(path = '', search = '')

Lists the files and folders under `path`. Returns an array of objects with `name`, `path`, `size`, `type` (`'file'` or `'folder'`), `createdBy`, `created`, `lastModifiedBy`, `lastModified` and `id`.

#### spdl.properties(path)

Returns the Graph properties of a single file (size, timestamps, ids).

#### spdl.propertiesP(paths)

Returns properties for an array of paths, fetched in parallel.

#### spdl.versions(path)

Lists the stored versions of a file — SharePoint keeps version history for you, which is one of the advantages of this data lake.

## Binary files

#### spdl.downloadBinary(remotePath, localPath)

Downloads a file from SharePoint to the local data folder. Returns `localPath`.

#### spdl.uploadBinary(remotePath, localPath, headers = null)

Uploads a local file (full path, may be outside the data folder) to SharePoint. Returns `remotePath`.

## Advanced

#### spdl.get(uri)

Performs a raw Graph API GET with the `spdl` authentication. When the reply is a paged OData collection, all pages are collected and returned as one array. Use this to reach any Graph endpoint the wrappers don't cover.
