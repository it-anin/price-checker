# GRAB CSV MANAGER / row review and export concepts

Open `/design-previews/grab-manager/index.html` for the earlier full-workspace concepts. The current link beside GRAB CSV MANAGER opens the corrected post-upload popup concepts at `/design-previews/grab-popups/index.html` (see [popup designs](grab-popup-designs.md)). The gallery contains ten distinct layouts; it does not replace the live manager. Preview uploads and comparisons run locally in the browser against fictional baselines and never write to Supabase.

| No. | Design | Structure |
|---|---|---|
| 01 | Studio Ledger | File status table, row table and inline inspector |
| 02 | Branch Deck | Three large branch cards with a lower review area |
| 03 | Review Inspector | File navigation, row list, permanent detail sidebar |
| 04 | Revision Journal | Vertical per-file timeline and expandable row changes |
| 05 | Before / After | Two parallel full-column panels for the selected row |
| 06 | Review Board | Changed/unchanged file lanes and an export lane |
| 07 | Change Atlas | SKU-by-column change matrix with a detail panel |
| 08 | Sheet Studio | Spreadsheet cells, row highlights and bottom file tabs |
| 09 | Guided Review | Select files, review changes, then export |
| 10 | Export Dock | Review workspace beside a selectable export packet |

All previews use the same generated CSV validator as the live app. They accept any filename, require the approved stock columns and StoreID, and compare decoded cell values across all 16 columns. Stable ItemID matches rows, falling back to SKU only when ItemID is blank; duplicate keys block comparison. Reordering rows does not count as a data change. Each changed row includes original/new CSV record numbers, changed columns and before/after values. Added, removed and unchanged rows have distinct text labels.

The shared demo has six products per branch: SRC has three edited rows, one addition and one removal; SSS is unchanged; KKL has two edited rows. Summary counts derive from these actual demo CSVs. A CSV uploaded in a preview is compared with that branch's original demo baseline. The footer can download that baseline, allowing an unchanged-file test.

Exports are functional: per-branch CSV exports preserve the entire incoming CSV with one UTF-8 BOM under `GRAB_XXX.CSV`; report CSVs contain only changed cells; ZIP contains the selected branches' CSVs. Exporting does not accept/save a template in the live app. No LOCK/UNLOCK controls are present.

Run `node scripts/build-grab-manager-previews.mjs` after changing the live CSV validator or `concepts.json`. It rebuilds the preview ES module, ten HTML entrypoints and gallery HTML. The preview app and CSS are maintained in `public/design-previews/grab-manager/`. Browser checks should cover all ten layouts at desktop/mobile widths, gallery navigation, row selection/filtering, unchanged and invalid uploads, and CSV/report/ZIP downloads.

After saving 1280×850 browser screenshots as `images/01.jpg` through `images/10.jpg`, run `python scripts/package-grab-manager-previews.py` (Pillow and DejaVu Sans required) to rebuild the overview, ten-page PDF and offline ZIP. The package checks PDF page count and ZIP integrity.
