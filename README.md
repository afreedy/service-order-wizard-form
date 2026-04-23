# CORA Service Order Code App

React, TypeScript, Vite, and Power Apps Code Apps application for creating CORA Environment service orders backed by SharePoint lists.

The app gives field users a guided service order wizard and gives authorized administrators a small reference-data console for drivers, vehicles, and waste categories.

## What the app does

- Creates service orders through a five-step wizard: basic details, customer, assignment, proof, and review.
- Loads customer hierarchy, driver, vehicle, waste-category, and existing service-order data from Power Apps generated services.
- Supports predefined customer selection from `customer-area-data-clean-final`.
- Supports ad hoc customer entry when the customer is not in the reference list.
- Captures a customer signature in the browser.
- Accepts before and after proof photos, compresses them client-side, and queues the media payload in SharePoint.
- Helps users read scale photos with local OCR and an optional remote OCR endpoint before they confirm tonnage.
- Creates one `service orders` row per completed waste line and also mirrors waste details to `service order waste items`.
- Provides admin-only pages for viewing service orders and maintaining drivers, vehicles, and waste categories.

## Tech stack

- React 19
- TypeScript 5.9
- Vite 7
- Power Apps Code Apps packages:
  - `@microsoft/power-apps`
  - `@microsoft/power-apps-vite`
- SharePoint Online connector generated services

## Repository layout

```text
.
├── .power/                         Power Apps schemas and data-source metadata
├── public/                         Static public assets
├── src/
│   ├── App.tsx                     Main application, wizard, admin pages, and submit flow
│   ├── App.css                     Application styles
│   ├── main.tsx                    React entry point
│   ├── weightOcr.ts                Scale-photo preprocessing, OCR, and tonnage parsing
│   ├── assets/                     Bundled image assets
│   └── generated/                  Power Apps generated models and services
├── docs/
│   ├── ARCHITECTURE.md             Runtime flow and code organization
│   ├── DATA-MODEL.md               SharePoint lists and field usage
│   └── OPERATIONS.md               Setup, build, deployment, and troubleshooting
├── power.config.json               Power Apps Code Apps configuration
├── vite.config.ts                  Vite configuration with Power Apps plugin
└── package.json                    Scripts and dependencies
```

## Local development

Install dependencies:

```powershell
npm install
```

Start the Vite development server:

```powershell
npm run dev
```

Build the production bundle:

```powershell
npm run build
```

Run linting:

```powershell
npm run lint
```

Preview a production build locally:

```powershell
npm run preview
```

Optional scale OCR endpoint:

```powershell
$env:VITE_WEIGHT_OCR_ENDPOINT="https://your-ocr-endpoint.example/weight"
npm run dev
```

The browser sends only the selected scale crop to this endpoint. Keep provider API keys on the server side; do not expose cloud OCR secrets through Vite environment variables.

## Power Apps configuration

`power.config.json` configures the Code Apps runtime:

- `buildPath`: `./dist`
- `buildEntryPoint`: `index.html`
- `localAppUrl`: `http://localhost:5173`
- SharePoint connector data sources:
  - `vehicles`
  - `drivers`
  - `drivers1`
  - `service order waste items`
  - `service orders`
  - `service order proof queue`
  - `waste_categories`
  - `customer-area-data`
  - `customer-area-data-clean`
  - `customer-area-data-clean-final`

Generated schema and service files are derived from `.power/` metadata. Treat `src/generated/` and `.power/schemas/appschemas/dataSourcesInfo.ts` as generated code unless you are intentionally refreshing Power Apps connector metadata.

## Documentation

Start here:

- [Architecture](docs/ARCHITECTURE.md)
- [Data model](docs/DATA-MODEL.md)
- [Operations](docs/OPERATIONS.md)
- [Power Apps OCR integration](docs/POWER-APPS-OCR.md)

## Admin access

Admin pages are gated by both the current Power Apps user email and a session passcode. The email allow-list and passcode currently live in `src/App.tsx`:

- `ADMIN_EMAILS`
- `ADMIN_PASSCODE`
- `ADMIN_UNLOCK_STORAGE_KEY`

Move these values out of client-side code before treating the app as security-sensitive. The current gate is a convenience control for the UI, not a substitute for SharePoint permissions or server-side authorization.

## Generated-code rule

Do not hand-edit generated models and services in `src/generated/`. When SharePoint columns or connector definitions change, refresh the Power Apps schemas and regenerate the service layer so types, connector metadata, and runtime calls stay aligned.
