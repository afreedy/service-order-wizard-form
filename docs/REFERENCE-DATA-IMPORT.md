# Reference Data Import

Reference data imports are programmatic and idempotent. Do not edit SharePoint rows manually through the list UI for routine driver, vehicle, or customer maintenance.

## Runtime Import Hook

After an allowed admin user unlocks the admin area, the app exposes:

```js
window.coraReferenceData.upsertReferenceData(config)
```

The hook runs inside the Power Apps host and writes through the generated SharePoint services:

- `Drivers1Service` for `drivers1`
- `VehiclesService` for `vehicles`
- `Customer_area_data_clean_finalService` for `customer-area-data-clean-final`

It uses the signed-in app context. It does not contain tenant secrets, connection strings, or credentials.

## Input Format

Use `tools/reference-data-import.template.json` as the starting point.

```json
{
  "drivers": ["HISHAM", "SAMUDI"],
  "vehicles": ["YR7210 Z", "GBJ8799A"],
  "customers": [
    {
      "Title": "Customer name",
      "field_1": "Location",
      "field_2": "Sub-location",
      "field_3": "Additional hierarchy or customer type",
      "field_4": "Address",
      "field_5": "Status",
      "field_6": "",
      "field_7": "",
      "field_8": "",
      "field_9": ""
    }
  ]
}
```

For the provided files:

- `List of vehicle for JKD Operation.xlsx`: vehicle numbers were copied into the template `vehicles` array.
- `Mona-Lisa existing customer list - 27 Apr.xlsx`: map `Level 1` to `Title`, `Level 2` to `field_1`, `Level 3` to `field_2`, `CUSTOMER TYPE` to `field_3`, `Address` to `field_4`, and `Status` to `field_5`.

## Safety Behavior

The importer:

- trims all input fields
- matches rows case-insensitively
- creates missing rows
- updates existing rows only when values differ
- skips rows that are already current
- rejects blank required keys
- rejects duplicate rows inside the import file
- fails a row safely if multiple existing SharePoint rows already match the same key
- returns a report with `created`, `updated`, `skipped`, and `failed` entries

Customer rows match on `Title`, `field_1`, and `field_2`; the remaining fields are updated if they differ.

## Example

Open the app through Power Apps, unlock Admin, then run:

```js
const config = {
  drivers: [
    "HISHAM",
    "SAMUDI",
    "Mohd Rafiq",
    "Ai Yi Song",
    "SORHONA",
    "JAHIS",
    "IDROS",
    "ABG ROZAIMI"
  ]
}

await window.coraReferenceData.upsertReferenceData(config)
```

Review the returned report before importing larger vehicle or customer batches.
