# Studio popup theme

The live app uses the workspace's cream, brown and orange tokens for every popup. Success and error states retain muted green and red. Shared `StudioModal` markup and `.pc-modal` styles cover GRAB price review, price calculation, Excel import, product creation/editing and image tools, GM MME comparison, missing branch products, Product Master comparison, and Promaxx comparison.

Each centered popup uses a native modal dialog with a shared outline heading icon, close control, footer, font and table styles. Content scrolls inside the popup; wide tables scroll horizontally. Toolbars wrap on small screens, and product forms become one column. Escape runs the same close handler as the close button, preserving each screen's existing reset behavior. Focus remains inside the active dialog and returns to its opener after closing.

GRAB CSV MANAGER retains its Side Notice drawer, using the same workspace accent, borders and text colors. Its comparison, CSV and report exports are unchanged.

`notifyPopup()` queues styled notices and confirmations. Callers await dismissal before continuing. Category deletion proceeds only after explicit confirmation; Cancel, Escape and the close control resolve as false. The cancel button receives initial focus for this destructive confirmation. Notices replace browser-native alerts in report validation, processing and template uploads.

Browser validation should open every popup on desktop and mobile, exercise search/tabs and exports, check long tables and editor fields, and verify Escape, focus containment/restoration and confirmation cancellation. Use mocked API and storage responses for these checks to avoid changing production data.
