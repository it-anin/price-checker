PRICE CHECKER / GRAB CSV MANAGER / 10 DESIGN CONCEPTS

Run a local HTTP server in this directory, for example:
  python -m http.server 8080
Then open http://localhost:8080/index.html.

The gallery shows 10 different layouts, with desktop/mobile previews.
All use fictional data and make no requests to production systems.
Uploads compare in your browser with a six-product demo baseline per branch.
The footer downloads that baseline so you can upload an unchanged file to test.
Exports download real CSVs, change reports, or a ZIP of selected branches.

Source filename and extension are unrestricted; CSV contents and StoreID must
still match the approved stock template. Original ItemID matches rows across
versions. Only a missing ItemID falls back to SKU. Duplicate keys block review.

Fonts: Noto Sans Thai, distributed under fonts/LICENSE.txt.
Generated engine.js comes from the repository's lib/process-engine.ts.
