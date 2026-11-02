# Archived design previews

The `tools/` and `grab-popups/` collections are kept here for reference, outside Next.js's `public/` directory. They are not served by the production app, and the live tools view has no links to them.

To view a collection locally, run one of these commands from the repository root and open `http://localhost:8080`:

```bash
python -m http.server 8080 --directory design-previews/tools
python -m http.server 8080 --directory design-previews/grab-popups
```

The live tools view uses Studio Cards, and GRAB CSV MANAGER uses the Side Notice popup. See [tools](../docs/tools-studio-cards.md) and [popup designs](../docs/grab-popup-designs.md) for details.
