# Operations

Use this guide to run, validate, and troubleshoot the CORA Service Order Code App.

## Prerequisites

- Node.js compatible with the installed Vite and TypeScript versions.
- Access to the target Power Apps environment.
- Access to the SharePoint site and lists configured in `power.config.json`.
- Power Apps connector permissions for the connected SharePoint data sources.

## Common commands

Install packages:

```powershell
npm install
```

Run locally:

```powershell
npm run dev
```

Build:

```powershell
npm run build
```

Lint:

```powershell
npm run lint
```

Preview the built app:

```powershell
npm run preview
```

## Local app URL

`power.config.json` sets the local app URL to:

```text
http://localhost:5173
```

Keep this aligned with the Vite dev server port when testing through Power Apps.

## Deployment checklist

Before publishing:

- Run `npm run lint`.
- Run `npm run build`.
- Confirm `dist/index.html` exists after the build.
- Confirm `power.config.json` points at the correct Power Apps environment and SharePoint data sources.
- Confirm the active customer source, `customer-area-data-clean-final`, contains rows.
- Test the service order wizard with one waste line and multiple waste lines.
- Test the proof step with a signature only, then with signature plus before/after photos.
- Test the admin pages with an allowed admin account.
- Confirm downstream proof processing handles new `service order proof queue` rows.

## SharePoint and connector changes

When a SharePoint list schema changes:

1. Refresh the Power Apps connector metadata.
2. Regenerate Power Apps schemas and TypeScript services.
3. Review changes in `.power/schemas/`.
4. Review generated model changes in `src/generated/models/`.
5. Update app code and documentation for any renamed, added, or removed fields.
6. Build and test the affected user flows.

Avoid editing generated files directly. Manual changes will be overwritten by regeneration and may hide schema drift.

## Troubleshooting

### Reference data does not load

Symptoms:

- The wizard shows `Reference Data Unavailable`.
- The retry button keeps returning an error.
- The error mentions `customer-area-data-clean-final`.

Checks:

- Confirm the SharePoint connector is available in the Power Apps environment.
- Confirm the current user has read access to all connected reference lists.
- Confirm `customer-area-data-clean-final` has at least one row.
- Confirm the data-source name in `power.config.json` matches the generated service name used by the app.
- Check browser console output for connector errors.

### Admin pages are hidden

Checks:

- Confirm `getContext()` returns the expected user principal name.
- Confirm the lower-cased user email is in `ADMIN_EMAILS` in `src/App.tsx`.
- Confirm the passcode was entered for the current browser session.
- Confirm the user has SharePoint permissions for the admin reference lists.

### Service order submit fails

Checks:

- Confirm `service orders`, `service order waste items`, and `service order proof queue` are writable for the user.
- Confirm all completed waste lines have both a category and numeric tonnage.
- Confirm the signature is present.
- If the error says records were rolled back, the app attempted to delete created rows after a later submit step failed.
- If the error says manual cleanup is needed, inspect the displayed service order IDs and proof queue ID.

### Photo upload fails

The app resizes photos to a maximum side length of 720 pixels and retries JPEG quality down to 0.32. If the base64 payload is still too large, users must choose a smaller photo.

Checks:

- Try a smaller image.
- Confirm the file is an image type.
- Confirm `service order proof queue` accepts the base64 field sizes currently being written.

### Scale OCR is inaccurate

Scale OCR uses the selected crop, local Tesseract passes, and an optional remote OCR endpoint. If `VITE_WEIGHT_OCR_ENDPOINT` is set, the browser posts the crop image to that endpoint and combines the returned text with local OCR evidence. Keep the OCR service key on the server side; do not put cloud OCR credentials in Vite environment variables.

Expected remote endpoint response shape:

```json
{
  "lines": [
    { "text": "1230 kg", "confidence": 0.94 }
  ]
}
```

Checks:

- Crop tightly around the scale display, including the unit when visible.
- Confirm detected values are reviewed with the `Use detected value` button; OCR does not write tonnage automatically.
- Confirm the returned value is within the app's default plausible display range of greater than 0 and up to 100000.
- If no unit is recognized, the app can still offer a detected value when clean OCR passes agree, but the user must confirm it with `Use detected value`.

### Proof URLs stay empty

The app queues proof media but does not itself upload files to storage. A downstream processor must:

- Find unprocessed queue rows.
- Decode and store `SignatureBase64`, `BeforePhotoBase64`, and `AfterPhotoBase64`.
- Write `SignatureUrl`, `BeforePhotoUrl`, and `AfterPhotoUrl`.
- Mark `Processed` as true or write `ErrorMessage`.

If proof URL fields stay empty, inspect the downstream Power Automate flow or processor rather than the React submit flow first.

## Operational risks

- Admin passcode and email allow-list are client-side values. Use platform permissions for real access control.
- The main app is concentrated in `src/App.tsx`, so feature changes can easily conflict. Prefer extracting focused modules during larger changes.
- Generated files must stay aligned with SharePoint schema changes.
- The proof queue stores base64 payloads in SharePoint fields; large images can exceed connector or list limits.
