# Data Model

The app uses SharePoint Online lists through the Power Apps connector. Connector metadata is stored under `.power/schemas/`, and TypeScript models/services are generated under `src/generated/`.

## Connected data sources

| Data source                      | Purpose                              | App usage                                                         |
| -------------------------------- | ------------------------------------ | ----------------------------------------------------------------- |
| `customer-area-data-clean-final` | Customer hierarchy reference data    | Drives the customer selection cascade in the service order wizard |
| `drivers1`                       | Driver reference list                | Driver dropdown and admin driver maintenance                      |
| `vehicles`                       | Vehicle reference list               | Vehicle dropdown and admin vehicle maintenance                    |
| `waste_categories`               | Waste category reference list        | Waste-category dropdown and admin category maintenance            |
| `service orders`                 | Primary service order records        | Created during submit and displayed in the admin order list       |
| `service order waste items`      | Waste line details                   | Created for each completed waste line                             |
| `service order proof queue`      | Signature and photo processing queue | Receives base64 proof payloads for downstream processing          |
| `drivers`                        | Legacy or alternate driver source    | Connected but not used by the current UI                          |
| `customer-area-data`             | Legacy or source customer data       | Connected but not used by the current UI                          |
| `customer-area-data-clean`       | Cleaned customer source              | Connected but not used by the current UI                          |

## Customer hierarchy

The active customer source is `customer-area-data-clean-final`.

Important generated fields:

| Field                       | Meaning in app                             |
| --------------------------- | ------------------------------------------ |
| `Title`                     | Customer name                              |
| `field_1`                   | Customer location or area fallback         |
| `field_2`                   | Sub-location fallback                      |
| `field_3` through `field_9` | Additional hierarchy-level fallback fields |

`App.tsx` also checks several friendly names dynamically when deriving customer hierarchy values, including `Area`, `SubLocation`, `Tower`, `LocationLevelN`, `LevelN`, and `CustomerLevelN`. This lets the app tolerate schema naming differences between cleaned customer lists.

The app exposes up to ten customer levels:

- `CustomerName`
- `CustomerLocation`
- `CustomerTenant`
- `CustomerLevel4`
- `CustomerLevel5`
- `CustomerLevel6`
- `CustomerLevel7`
- `CustomerLevel8`
- `CustomerLevel9`
- `CustomerLevel10`

## Service orders

Primary list: `service orders`

Fields used by the app:

| Field                                      | Type    | Usage                                         |
| ------------------------------------------ | ------- | --------------------------------------------- |
| `Title`                                    | string  | Generated service order number                |
| `Customer`                                 | string  | Customer display value                        |
| `CustomerName`                             | string  | Selected or ad hoc customer name              |
| `CustomerLocation`                         | string  | Selected or ad hoc customer location          |
| `CustomerTenant`                           | string  | Selected or ad hoc sub-location               |
| `CustomerLevel4` through `CustomerLevel10` | string  | Additional selected customer hierarchy levels |
| `IsAdhocCustomer`                          | boolean | Marks manually entered customer details       |
| `DriverName`                               | string  | Selected driver                               |
| `VehicleNumber`                            | string  | Selected vehicle                              |
| `DateOfCollection`                         | string  | Collection date from the wizard               |
| `WasteCategory`                            | string  | Waste category for this row                   |
| `Tonnage`                                  | number  | Tonnage for this row                          |
| `SignatureUrl`                             | string  | Signature output URL when available           |
| `BeforePhotoUrl`                           | string  | Before photo output URL when available        |
| `AfterPhotoUrl`                            | string  | After photo output URL when available         |
| `Notes`                                    | string  | Free-text notes                               |

When a user enters multiple completed waste lines, the app creates multiple `service orders` rows with the same base order information and different `WasteCategory`/`Tonnage` values.

Scale OCR is a capture aid only. Scale photo files, crop rectangles, OCR text, OCR confidence, and OCR reasons stay in browser state and are not written to SharePoint. Only the user-confirmed `Tonnage` value is persisted.

## Waste items

Primary list: `service order waste items`

Fields used by the app:

| Field           | Usage                    |
| --------------- | ------------------------ |
| `Title`         | Service order title      |
| `WasteCategory` | Completed waste category |
| `Tonnage`       | Completed tonnage        |

The app writes these rows after service-order rows are confirmed. They provide a normalized companion record for each completed waste line.

## Proof queue

Primary list: `service order proof queue`

Fields used by the app:

| Field                 | Usage                                          |
| --------------------- | ---------------------------------------------- |
| `Title`               | Queue item title, derived from the order title |
| `ServiceOrderId`      | First created service-order row ID             |
| `OrderTitle`          | Service order title                            |
| `SignatureFileName`   | Signature output file name                     |
| `SignatureBase64`     | Signature image payload                        |
| `BeforePhotoFileName` | Optional before photo output file name         |
| `BeforePhotoBase64`   | Optional before photo payload                  |
| `AfterPhotoFileName`  | Optional after photo output file name          |
| `AfterPhotoBase64`    | Optional after photo payload                   |
| `Processed`           | Set to `false` by the app                      |
| `ErrorMessage`        | Available for downstream processors            |
| `SignatureUrl`        | Available for downstream processors            |
| `BeforePhotoUrl`      | Available for downstream processors            |
| `AfterPhotoUrl`       | Available for downstream processors            |

The app queues proof data only. A separate Power Automate flow or backend processor should read unprocessed queue rows, persist media to the desired storage location, write output URLs, and mark the queue row as processed.

## Reference lists

`drivers1`, `vehicles`, and `waste_categories` currently use the SharePoint `Title` field as their display value. Admin pages create, update, and delete these records through generated services.

## Generated files

Do not hand-edit:

- `src/generated/models/*.ts`
- `src/generated/services/*.ts`
- `.power/schemas/appschemas/dataSourcesInfo.ts`

Refresh or regenerate them after connector or SharePoint schema changes.
