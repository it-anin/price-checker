# PENDING SKU UPDATE

The PENDING card in PROCESS ENGINE manages products waiting for Grab approval. Read-only visitors can inspect entries and compare prices. ADD/DELETE use the same Supabase Auth session as GRAB CSV MANAGER and do not require its UNLOCK switch.

## Managing pending entries

Select SRC, SSS and/or KKL (all selected initially), then paste SKUs separated by newlines, commas, semicolons, tabs or pipes. Enter submits; Shift+Enter adds a line. Existing SKUs gain the selected branches without duplicates. DELETE removes only the active tab's branch and removes the entry only when no branches remain. WRD/ONN memberships and additional stored metadata are preserved; the total badge includes these legacy records.

New entries save before optional name lookup. Names come from approved stock templates, using their actual headers (C=ITEMNAME in current stock exports, SKU D). Lookup stops when all missing names are found and caches each file by name/updated_at. Name lookup never changes the uploaded GRAB price maps.

## Three price sources

- **GRAB:** independent CSV uploads per branch, named GRAB_SRC.CSV / GRAB_SSS.CSV / GRAB_KKL.CSV (case-insensitive). Headers must identify name B, price C and **SKU I**. StoreID is validated when present. A valid replacement replaces that branch's whole price map; a rejected file preserves the last successful map. Stock templates are not price files.
- **DATABASE:** current `products` selling price `*ราคาสินค้า`. Uses the existing read-only `POST /api/products` SKU lookup in parallel batches of at most 300. Loads on opening/reloading the card, after ADD/DELETE, and when RUN refreshes the pending snapshot. REFRESH PRICE affects only this source. Failed refreshes keep the previous prices with an error message.
- **PROMAX:** R05.103 CSV, SKU B / filter E=`5` / price G. Only matching filter rows are included. Uploads support UTF-8 and TIS-620 and remain in memory.

Prices use the existing `parsePriceRobust` and compare rounded cents. `·` means the source is not loaded; `—` means no valid price for the SKU; zero displays `0.00`. GRAB differences against either reference are marked with a red row and `≠`, with separate per-tab counts. CLEAR affects only its source. Price uploads never write product prices or branch availability. In-memory sources survive workspace navigation but disappear on page refresh.

## Safe persistence and RUN

The existing `app_settings` row `key='pending_skus'` remains the source of truth. `loadPendingSnapshot()` returns normalized entries plus `updatedAt`. `savePendingMutation()` handles `add`, `removeBranch` and `enrichNames`: it re-reads current data and updates with an `updated_at` equality filter, advancing the timestamp beyond its prior value. On a version conflict it reloads/reapplies the same operation, up to three attempts. Every writer in this UI uses this path; external/manual writers should also advance `updated_at`.

Failed loads, malformed data and failed saves block subsequent edits until a successful reload. Input is retained when ADD fails. Membership edits invalidate previously generated stock outputs so they must be generated again; this also applies when an explicit reload observes another user's changes. Name/metadata-only updates do not invalidate stock results. Writes are serialized against RUN and template uploads.

RUN always reads current pending data, applies suppression only to the SKU's branch, and stops if the source cannot be read. Authenticated runs may fill missing names using the same mutation helper; a name-save failure is logged and does not discard otherwise valid stock outputs. If name enrichment observes changed active membership during RUN, outputs are withheld and the user is asked to run again. No schema migration or new API route is required.

## Verification

```sh
node --test tests/process-engine.test.cjs tests/pending-sku.test.cjs
npm run build
```

Tests cover branch merging/removal, legacy metadata, CAS retries/failures, invalid sources, timestamp monotonicity, column-I price parsing, actual stock-template names, TIS-620, cent comparisons, and stock suppression after ADD/DELETE. Browser QA exercises the UI with mocked API responses and performs no production writes.
