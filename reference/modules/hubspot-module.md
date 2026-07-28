---
description: Read and write HubSpot CRM objects
---

# Hubspot module

Wrappers around the [HubSpot CRM API](https://developers.hubspot.com/docs/api/crm/understanding-the-crm) (`auth = 'hubspot'`). Read wrappers handle paging automatically and return one flat array.

```javascript
hubspot = require('hubspot.js')

deals = hubspot.deals()
```

## Reading

* `hubspot.deals(args)`
* `hubspot.companies(args)`
* `hubspot.contacts(args)`
* `hubspot.line_items(args)`
* `hubspot.quotes(args)`
* `hubspot.schemas(args)`

The optional `args` works like this:

1. **Omitted** — all records are returned (paged in batches of 100).
2. **A single value** — appended to the uri as a path, returning one record: `hubspot.deals('12345')`
3. **An object** — appended as query parameters: `hubspot.deals({ properties : 'dealname,amount' })`

## Writing

#### hubspot.patchDeal(dealId, properties)

Updates properties on a deal:

```javascript
hubspot.patchDeal('12345', { dealstage : 'closedwon' })
```

#### hubspot.writeDealLineItems(dealId, lines)

Creates line items and associates them with a deal. Each line is an object of [line item properties](https://developers.hubspot.com/docs/api/crm/line-items); the position on the quote is set from the array order:

```javascript
lines = [
    { name : 'Consulting', price : 120.00, quantity : 10 },
    { name : 'Support',    price : 80.00,  quantity : 2 }
]
hubspot.writeDealLineItems('12345', lines)
```

#### hubspot.deleteDealLineItems(dealId)

Deletes all line items associated with a deal. Combine with `writeDealLineItems` to replace a deal's lines.

## Advanced

#### hubspot.callAPi(api, args)

Calls any HubSpot api by its relative path with the same `args` convention and paging as the wrappers:

```javascript
tickets = hubspot.callAPi('crm/v3/objects/tickets')
```

#### hubspot.spec()

Returns HubSpot's public API catalog, so you can browse the available apis from Javascript Studio.
