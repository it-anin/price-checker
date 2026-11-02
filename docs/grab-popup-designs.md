# Post-upload change popup designs

The current GRAB CSV MANAGER preview link opens `/design-previews/grab-popups/index.html`. These are ten **popup designs that appear after an upload when file data changed**, each over the same manager background. The previous full-workspace concepts remain at `/design-previews/grab-manager/index.html`.

| No. | Popup | Structure |
|---|---|---|
| 01 | Soft Summary | Compact centered warning and three row counters |
| 02 | Diff Table | Wide modal with a cell-change table |
| 03 | Before / After | Modal comparing the selected row in two panels |
| 04 | Side Notice | Modal drawer at the right edge, with dimmed backdrop |
| 05 | Bottom Review | Bottom sheet with horizontally arranged change cards |
| 06 | File Receipt | Narrow receipt with dashed separators and aligned counts |
| 07 | Impact Split | Dark summary rail and a separate changed-row area |
| 08 | Change Timeline | Edits, additions and removals grouped on a vertical line |
| 09 | Branch Tabs | One batch-upload popup with per-branch result tabs |
| 10 | Searchable Review | Searchable changed-row list and selected-row details |

Open a concept with `?demo=1` to show the popup immediately in the simulated post-upload state. Without that query, the popup starts closed. “ลองอัปไฟล์ที่มีข้อมูลเปลี่ยน” simulates one changed upload; “ลองอัป 3 ไฟล์พร้อมกัน” simulates a batch. A changed upload opens the popup automatically. An unchanged upload shows a non-modal status and no change popup; invalid CSVs show an inline error and no change popup. Actual uploads are compared locally against fictional branch baselines. No production data is fetched or written, and the live manager upload flow is not replaced by a selected concept yet.

All variants use native `dialog.showModal()` with a dimmed backdrop, accessible heading/description, focus containment, Escape dismissal and restored opener focus. Each has close/acknowledge controls plus working CSV and change-report exports. Batch tabs in 09 also show unchanged branches; the other variants show changed branches only.

Run `node scripts/build-grab-popup-previews.mjs` to generate entrypoints, the gallery and shared CSV modules. Save 1280×850 screenshots as `images/01.jpg` through `images/10.jpg`, then run `python scripts/package-grab-manager-previews.py --collection grab-popups` for the overview, PDF and ZIP.
