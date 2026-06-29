# Version Bumping Rule

Every time a code change is made to this project, you MUST increment the version number in `app/page.tsx`.

## Current Version: v1.11

## Location
```
app/page.tsx — line with:
<h1 ...>Grab Master รายการสินค้าทั้งหมด by mailforspiritwish <span ...>vX.X</span></h1>
```

## Rules
- Version format: `v{major}.{minor}` — increment minor by 1 for each change session
- Current: **v1.9** → next change → **v1.10** → **v1.11** → etc.
- Always update this file's "Current Version" line after bumping
- Include the version bump in the same commit as the code change
