---
description: Pull accounting data from Xero, with local sync for fast repeated pulls
---

# Xero module

Wrappers around the [Xero API](https://developer.xero.com/documentation/api/accounting/overview) (`auth = 'xero'`). Xero uses one login for all your organisations — every call takes a `tenantId` to say which organisation to query.

The heavier data sets (journals, invoices, budgets) are **synced incrementally to the local data folder** (under `xero/{tenantId}/...`): the first pull downloads everything, subsequent pulls only fetch what changed since. This makes repeated reporting pulls fast and light on Xero's rate limits.

```javascript
xero = require('xero.js')

conns = xero.connections()          // list your organisations
tenantId = conns[0].tenantId

accounts = xero.accounts(tenantId)  // chart of accounts
```

## Organisations

#### xero.connections()

Lists the Xero organisations connected to your login, sorted by name. Each row has a `tenantId`.

#### xero.Organisation(tenantId)

Returns the organisation details (name, short code, financial year end, ...).

#### xero.getMarkedTenantId(Connections) / xero.getMarkedTenantIds(Connections)

For the common workbook pattern where a `Connections` table is imported into Excel with a `Pull` checkbox column: returns the tenantId(s) of the marked row(s), with validation. The single version throws unless exactly one row is marked.

## Master data

#### xero.accounts(tenantId, addComputedAccounts = false)

Chart of accounts. With `addComputedAccounts` true, two virtual accounts are appended that Xero reports use but the Accounts api doesn't return: *Unrealised Currency Gains* and *Current Year Earnings*.

#### xero.taxRates(tenantId)

#### xero.trackingCategories(tenantId)

## Reports

#### xero.profitAndLoss(tenantId, tenantName, fromDate, toDate, accountingBasis = 'Accrual', tc1Id = null, tc2Id = null)

Pulls a P\&L report and unpacks Xero's nested report format into a flat table of `{ AccountID, Period, TC1, TC2, Value }` rows (zero rows dropped). Pass tracking category ids to break down by tracking category; `accountingBasis` can be `'Accrual'` or `'Cash'`.

#### xero.balanceSheet(tenantId, period, accountingBasis = 'Accrual')

Pulls a balance sheet for `period` (yyyy-mm-dd) and returns flat `{ AccountID, Period, Value }` rows.

#### xero.trialBalance(tenantId, date, paymentsOnly = false)

Returns flat rows with `Account`, `AccountID`, `Period`, `Debit`, `Credit`, `YTDDebit`, `YTDCredit`.

#### xero.bankTransactions(tenantId, fromDate = null, toDate = null)

Bank transactions, filtered locally by date when arguments are given.

## Journals

#### xero.syncJournals(tenantId, tenantName, accBasis)

Incrementally downloads new journals since the last sync into monthly files in the local data folder. `tenantName` is only used for progress messages; `accBasis` is `'Accrual'` or `'Cash'`.

#### xero.pullJournals(tenantId, tenantName, accBasis, startDate, endDate, syncFirst = true)

Returns all journal **lines** between the dates as one flat array — each row is the journal header merged with the line, with `TrackingCategory1`/`TrackingCategory2` resolved to option names and dates converted to ISO format. Syncs first by default.

#### xero.periodsPLBSFromJournals(Connections, endDate, numPeriods, accBasis = 'ACCRUAL', syncFirst = true, skipFXGROUPIDCheck = false, returnTS = true)

The one-call engine behind multi-company reporting workbooks. Takes a `Connections` table (or a single tenantId string) and returns monthly P\&L figures computed from the synced journals, for `numPeriods` months ending at `endDate` (yyyy-mm-dd):

```javascript
res = xero.periodsPLBSFromJournals(Connections, '2026-06-30', 12)

res.PL   // P&L rows: AccountID, Period, TC1, TC2, Value, account details, Company
res.TS   // the underlying journal lines (transactions), incl. DeepLink to Xero
res.ACC  // chart of accounts per company
res.ORG  // organisation details per company
```

Throws when the journals contain multiple currencies (not supported yet) unless `skipFXGROUPIDCheck` is true. Set `returnTS` false to skip returning the raw transactions.

## Invoices

Invoice-type data is synced locally per type: `Invoices`, `CreditNotes`, `Overpayments`, `Prepayments`.

#### xero.syncInvoiceType(tenantId, fileType) / xero.syncAllInvoiceTypes(tenantId)

Incrementally sync one or all of the invoice types to the local data folder.

#### xero.pullInvoices(tenantId, invoiceType, asAtDate, syncFirst = true)

Returns the invoices **outstanding as at** `asAtDate` — totals and payments are reconstructed to that date, so you can produce a historical debtors/creditors position. `invoiceType` is `'ACCREC'` (sales) or `'ACCPAY'` (purchases). Rows include totals, paid and remaining amounts, contact and dates.

#### xero.pullCreditNotes(tenantId, asAtDate, syncFirst = true)

#### xero.pullOverpayments(tenantId, asAtDate, syncFirst = true)

#### xero.pullPrepayments(tenantId, asAtDate, syncFirst = true)

Same as-at treatment for the other outstanding-amount types.

## Budgets

#### xero.budgets(tenantId, fromDate, toDate)

Pulls all budgets and unpacks them into flat `{ Budget, Period, AccountID, AccountCode, Amount, TC1, TC2 }` rows.

#### xero.pullBudgets(tenantId, fromDate, toDate)

Like `budgets`, but cached in the local data folder in two-year blocks (Xero serves at most 24 months per call) — much faster for repeated pulls. `xero.syncBudgets(tenantId)` drops the cache for budgets that changed in Xero, and runs automatically as part of `pullBudgets`.

## Helpers

#### xero.periods(ToDate, Periods)

Generates a list of `Periods` month-ends going back from `ToDate`, as `{ fromDate, toDate, budgetPeriod }` — handy to drive period loops.

#### xero.parseXeroDate(xeroDate)

Converts Xero's `/Date(1234567890000)/` format to a javascript Date.

#### xero.xeroHeader(tenantId)

Returns `{ 'xero-tenant-id' : tenantId }`, the header needed for direct http calls.

#### xero.deepLink(ShortCode, SourceType, SourceID)

Builds a url that opens the underlying document (invoice, bank transaction, manual journal, ...) directly in Xero — great for making report rows clickable. The organisation `ShortCode` comes from `xero.Organisation()`.

#### xero.sourceLabel(sourceCode) / xero.typeLabel(typeCode)

Translate Xero's source codes (`ACCREC`, `MANJOURNAL`, ...) and account type codes (`CURRLIAB`, `OVERHEADS`, ...) to human-readable labels.
