---
description: Pull data from the Brincr API with automatic paging and rate limiting
---

# Brincr module

Helper for the [Brincr API](https://sandbox-api.brincr.com/api/documentation). Brincr uses a [Visma Client Credentials](../../developing-with-xlconnect/cloud-systems.md#visma-client-credentials-service) (Service) authentication.

#### brincr.getAllPages(uriBase, auth, limit = 200)

Fetches every page of a paged Brincr api and returns one flat array. The function follows the `links.next` uri in each reply, and watches the `X-RateLimit-Remaining` header — when the rate limit is nearly used up it pauses for 5 seconds (showing a progress message) before continuing.

```javascript
brincr = require('brincr.js')

uri = 'https://api.brincr.com/api/v1/products'
products = brincr.getAllPages(uri, 'vismacc')
```
