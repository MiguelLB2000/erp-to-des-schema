# erp-to-des-schema

Converts an ERP routing export into a populated **Discrete Event Simulation (DES) input schema** workbook.

The tool reads raw operation/routing data, decodes cycle times according to configurable rate-type rules, deduplicates to the latest routing version, and writes Routing, Products, and Machines sheets into a schema template — without touching any Python when you switch ERP sources.

---

## What it does

| Step | Action |
|------|--------|
| Load | Reads the ERP export (`.xls` or `.xlsx`) |
| Filter | Keeps only the target plant/site; drops admin-note rows and cost-type rate rows |
| Deduplicate | Retains the highest routing-plan version per part + operation |
| Decode | Converts ERP rate values → `cycle_time_min` using per-ZBS-code formulas from config |
| Write | Populates `Routing`, `Products`, and `Machines` sheets in the schema template |
| Report | Prints a validation summary (rows processed/skipped, TODO fields, resource groups) |

---

## Install

```bash
pip install -r requirements.txt
```

Dependencies: `pandas`, `openpyxl`, `xlrd`, `PyYAML`

---

## Usage

```bash
python extract.py \
    --input  alle_AGs.xls \
    --config config.yaml \
    --schema DES_Input_Schema_v04.xlsx \
    --output Factory1_Schema_v04.xlsx
```

| Flag | Required | Description |
|------|----------|-------------|
| `--input` | yes | ERP routing export (`.xls` or `.xlsx`) |
| `--config` | yes | YAML config file (column names, ZBS rules) |
| `--schema` | yes | DES schema template workbook (provides sheet structure) |
| `--output` | yes | Path for the populated output workbook |

---

## Configuring `config.yaml` for a new ERP source

All ERP-specific knowledge lives in `config.yaml`. You never need to edit `extract.py`.

### 1 — Map your column names

Open your ERP export and find the exact header text for each logical field:

```yaml
columns:
  part_id:              "Teil"          # part / product number
  part_name:            "Bezeichnung"   # part description
  routing_step:         "Arbeitsgang"   # operation sequence number
  routing_plan:         "Arbeitsplan"
  routing_plan_version: "Arbeitsplan-Version"
  operation_name:       "AG-Bezeichnung"
  cost_center:          "Kst."          # maps to resource group (RG_xxx)
  machine_id:           "Fert.anlage"   # individual machine ID
  calc_rate:            "kalk. Vorgabe" # numeric rate value
  zbs:                  "ZBS"           # rate-type code
  site:                 "Standort"      # plant / site filter column
```

Replace the quoted strings with whatever your ERP export uses.

### 2 — Set the site filter

```yaml
site_filter:
  value: "OV1"   # keep only rows where columns.site equals this
```

### 3 — Define ZBS / rate-type decode rules

`skip_values` lists codes that mean "not a time row" (cost rates, blanks). These rows are dropped entirely.

`conversions` maps each rate-type code to a formula:
`cycle_time_min = (rate_value / divisor) × multiplier`

```yaml
zbs_rules:
  skip_values:
    - ""
    - "C"         # Kosten pro Stück — cost row, not time
  conversions:
    "3":
      description: "hours per 1 000 pieces"
      divisor: 1000
      multiplier: 60
    "2":
      description: "hours per 100 pieces"
      divisor: 100
      multiplier: 60
```

For a SAP PP export where the rate is already in minutes per piece:

```yaml
  conversions:
    "MIN":
      description: "minutes per piece"
      divisor: 1
      multiplier: 1
```

### 4 — Skip admin-note operation names

```yaml
skip_operation_names:
  - "Hinweis"   # add any operation labels that are notes, not real steps
```

### 5 — Default output values

Fields with no ERP source column receive these defaults. Adjust as needed:

```yaml
defaults:
  batch_size: 1
  yield_pct:  100.0
  # … see config.yaml for full list
```

See [`examples/factory2_config.yaml`](examples/factory2_config.yaml) for a SAP PP stub showing a different column-name and rate-unit mapping.

---

## Output schema

The output workbook contains three sheets populated from the ERP data:

### Routing
One row per part × operation step.

| Column | Source | Notes |
|--------|--------|-------|
| `product_id` | `Teil` | |
| `step_number` | `Arbeitsgang` | integer |
| `process_name` | `AG-Bezeichnung` | |
| `resource_group` | derived from `Kst.` | `RG_<kst>` — rename to descriptive label |
| `cycle_time_min` | decoded from `kalk. Vorgabe` + `ZBS` | |
| `setup_time_min` | default 0.0 | fill from setup-time source if available |
| `batch_size` | default 1 | fill from PB-Familien.xlsx |
| `yield_pct` | default 100 | fill from scrap data |
| other columns | defaults from config | |

### Products
One stub row per unique part number. `lot_size` and `arrival_interval_min` are written as `TODO` — fill from demand/sales data.

### Machines
One stub row per unique `Fert.anlage` × resource group. `MTBF_hours` and `MTTR_min` are `TODO` — fill from maintenance data (e.g. Stammdaten.xlsx).

---

## Validation summary

After every run the tool prints:

```
──────────────────────────────────────────────────────
  VALIDATION SUMMARY
──────────────────────────────────────────────────────
  Rows loaded from ERP file          59,574
  Rows written to Routing sheet      57,312
  Rows skipped (all reasons)          2,262
──────────────────────────────────────────────────────
  Skip reasons:
    ZBS_skip (cost/blank)              1,686
    operation='Hinweis'                1,316
    wrong_site (≠OV1)                      0
    invalid_step_number                  ...
  Fields populated                        14
  Fields still TODO                        3
    → MTBF_hours
    → MTTR_min
    → lot_size
  Unique resource groups                  12
  Unique products                      4,821
  Machine stubs written                  237
──────────────────────────────────────────────────────
```

---

## Repository layout

```
erp-to-des-schema/
├── README.md
├── config.yaml               default config (Factory 1 column names)
├── extract.py                main script
├── requirements.txt
└── examples/
    ├── factory1_config.yaml  identical to default — reference copy
    └── factory2_config.yaml  stub for SAP PP routing export
```

---

## Next steps after extraction

1. **Rename resource groups** — open the output workbook, find the `RG_xxx` codes in the Routing sheet, and replace them with descriptive names (e.g. `RG_CASTING`, `RG_CORE`) using your Kostenstellen reference.
2. **MTBF / MTTR** — fill the Machines sheet from maintenance data.
3. **Lot sizes** — fill `lot_size` in the Products sheet from batch-family data.
4. **Arrival rates** — fill `arrival_interval_min` from sales/demand data.
5. **Shifts** — populate the Shifts sheet from your shift schedule.
