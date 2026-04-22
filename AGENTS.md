# AGENTS.md

This file gives coding agents the minimum project context needed to work safely in this repository.

## Project Summary

- App: CORA Service Order Code App
- Stack: React 19, TypeScript 5.9, Vite 7, Power Apps Code Apps
- Runtime: Browser SPA hosted through Power Apps, with generated SharePoint connector services
- Main purpose: create service orders, queue proof media, and maintain admin reference data

## Important Paths

- `src/App.tsx`: primary application file; contains most wizard, admin, state, and submit logic
- `src/weightOcr.ts`: browser OCR preprocessing and tonnage parsing
- `src/generated/`: generated Power Apps models and services
- `.power/`: Power Apps schemas and connector metadata
- `docs/ARCHITECTURE.md`: runtime flow and code organization notes
- `docs/DATA-MODEL.md`: SharePoint list and field mapping
- `docs/OPERATIONS.md`: setup, build, deployment, and troubleshooting
- `power.config.json`: Power Apps configuration and data-source definitions

## Common Commands

Run from the repository root:

```powershell
npm install
npm run dev
npm run build
npm run lint
npm run preview
```

Optional remote OCR endpoint during local development:

```powershell
$env:VITE_WEIGHT_OCR_ENDPOINT="https://your-ocr-endpoint.example/weight"
npm run dev
```

## Working Rules

- Read `README.md` and the docs in `docs/` before making non-trivial changes.
- Prefer targeted edits; `src/App.tsx` is large and easy to destabilize.
- Preserve the current confirm-before-write OCR flow. OCR should suggest values, not silently overwrite tonnage.
- Keep admin gating behavior consistent unless the task is explicitly about auth/security.
- When changing submit logic, be careful with rollback behavior for service orders, waste items, and proof queue records.
- Validate both the end-user wizard flow and admin maintenance flow after meaningful UI or data changes.

## Generated Code

Treat these as generated unless the task is explicitly to refresh connector output:

- `src/generated/**`
- `.power/schemas/**`
- `.power/schemas/appschemas/dataSourcesInfo.ts`

Do not hand-edit generated files as a normal feature fix. If SharePoint schema or connector metadata changes, regenerate first, then update app code around the regenerated output.

## Suggested Change Strategy

For small fixes:

- Keep changes localized.
- Run `npm run lint`.
- Run `npm run build` if types, imports, or data flow changed.

For larger feature work:

- Extract stable logic from `src/App.tsx` into focused modules or components instead of growing the file further.
- Favor folders like `components/`, `features/`, or `lib/` when introducing reusable code.
- Update documentation when behavior, data contracts, or operational steps change.

## Testing Expectations

There is no dedicated automated test suite in the repo today, so verification is mainly:

- `npm run lint`
- `npm run build`
- manual checks of the affected wizard/admin flow

If you cannot run one of these checks, say so clearly in your handoff.

## Operational Notes

- Reference data loading depends on `customer-area-data-clean-final` having rows.
- Admin access currently depends on client-side values in `src/App.tsx` (`ADMIN_EMAILS`, `ADMIN_PASSCODE`, `ADMIN_UNLOCK_STORAGE_KEY`).
- Proof media is queued in SharePoint; the React app does not upload files directly to final storage.
- Large images may fail if the proof queue payload exceeds connector or SharePoint limits.

## Handoff Checklist

Before finishing, note:

- what changed
- whether `npm run lint` passed
- whether `npm run build` passed
- any manual validation performed
- any follow-up risk, especially around generated services, SharePoint schema assumptions, or submit rollback behavior
