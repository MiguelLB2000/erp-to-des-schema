# erp-to-des-schema

Maps ERP or MES exports into the input workbook of a generic discrete-event simulation (DES) model. Headers may be German or English. Everything that is certain is filled automatically; everything uncertain becomes a question; everything missing is listed, required columns first.

The target is `schema/Schema_Template_EMPTY.xlsx`: one sheet per input table (Products, Routing, Machines, Shifts, Workers, Orders, BOM, ...) and a `Dictionary` sheet that states for every column whether it is required and what the model assumes when it is blank. The tool reads that Dictionary, so it follows the template and has no column rules of its own.

---

## Data protection

Plain Python that runs on your own computer. No network calls, no AI or API calls (`map_schema.py` imports no HTTP or model library). Your exports never leave the machine. Results are written next to the output path you give. The `.gitignore` blocks every data and workbook type, so exports, generated workbooks, answers files and reports cannot be committed by accident.

---

## Install

```bash
pip install -r requirements.txt
```

Reading `.accdb` / `.mdb` needs Windows with the Microsoft Access ODBC driver (included with Office, or the free "Access Database Engine"). Without it the tool says so; export the tables to `.xlsx` or `.csv` instead.

---

## Usage

```bash
python map_schema.py --input <file or folder> --output my_factory.xlsx
```

| Flag | Meaning |
|------|---------|
| `--input` | One export (`.csv`, `.txt`, `.xls`, `.xlsx`, `.xlsm`, `.accdb`, `.mdb`) or a folder of them. Every sheet of a workbook and every table of an Access file is read. |
| `--output` | Workbook to write. |
| `--config` | Optional rules file: row filters, rate-code unit conversion, id prefixes, latest routing version (see below). |
| `--answers` | Answers of an earlier run. Default: `<output>_answers.yaml` if it exists. |
| `--no-input` | Ask nothing; open questions are listed in the report. |
| `--schema`, `--synonyms` | Other template or synonym list (defaults: `schema/Schema_Template_EMPTY.xlsx`, `synonyms.yaml`). |

Try it on the made-up demo (German headers, one workbook with two sheets plus two CSV files):

```bash
python map_schema.py --input examples/demo --output demo.xlsx
```

---

## What happens in a run

1. **Read.** Every table is loaded; the header row is found automatically (title rows above it are skipped). CSV separators `;` `,` tab `|` and encodings UTF-8 / Windows-1252 are detected.
2. **Tables to sheets.** Each table is scored against each template sheet by its headers. A table is assigned without a question only when at least two columns match with certainty and it clearly beats every other table. Otherwise you choose from a short list (or "none").
3. **Columns.** For each sheet column:
   - a header equal to a known synonym is assigned (`Artikelnummer` -> `product_id`, `Arbeitsgang` -> `step_number`, `Kostenstelle` -> `resource_group`);
   - a header that may mean this column or another one becomes a question (`Menge`, `Anzahl`, `resource_id`, `quantity_on_hand` vs `available_qty`);
   - columns still missing are listed per sheet, required ones first. You choose: **fill by hand later**, or **give the file and column**. A column from another table is linked by a column both tables share (for example `bom_id`); if a key appears several times there, the first row is used and the report says so.
4. **Units.** Time columns are converted to the template unit (`_min` or `_hours`). The unit is read from the header (`[min]`, `[h]`, `minutes`, `Stunden`); if the header does not say, you are asked. Date columns (release, due) become minutes after one simulation start, which is asked once.
5. **Write and report.** The workbook is written with English template headers; `Config scenario_id` is set to the file name. A report lists what is still empty: required gaps first, then optional ones with the default the model will use. Saved as `<output>_report.txt`.

Every answer is saved in `<output>_answers.yaml`. Run again and nothing already answered is asked; edit that file to change a choice. The answers file is also the written record of how this company's data was mapped.

**What the tool does not do:** guess meanings or units, compute derived values (sums per product, arrival intervals, capacities from several tables), clean data, or fill Config (run length, default shift, modes). Those stay with you; the report and the `Dictionary` sheet say what each column means.

---

## Teaching it your headers

`synonyms.yaml` holds, per template column, the headers that mean it for sure (`exact`) and the ones that might (`ambiguous`). Matching ignores case, umlauts, punctuation and text in brackets. Add your ERP's headers there; no code change is needed.

---

## Rules file (`--config`, optional)

```yaml
filters:                      # keep only rows where a column has this value
  make_buy: MAKE
skip_values:                  # drop rows where a column has one of these values
  operation_name: [note, blocked]
unit_rules:                   # convert a column by a rate-code column
  cycle_time_min:
    code_column: rate_code
    skip_codes: ["", "C"]     # rows with these codes are not time data
    codes:
      "3": {divisor: 1000, multiplier: 60}   # hours per 1,000 pieces -> minutes per piece
      "2": {divisor: 100,  multiplier: 60}
transforms:                   # id changes, per column or per sheet.column
  resource_group: {prefix: "RG_"}
  Alternatives.product_id: {strip_prefix: "RTG-"}
dedupe:                       # keep the highest version per key
  keys: [part, plan, step]
  version: plan_version
columns_map:                  # force a header for a column
  Routing.cycle_time_min: "Standard Value 1"
```

`config.yaml` is a complete example for an ERP routing export with rate codes (ZBS), a site filter, note rows and routing versions; the older format of `extract.py` is read as well.

---

## Tested on

| Data | Input | Result |
|------|-------|--------|
| Factory A, ERP routing export (German headers, rate codes, 59,258 rows) | one `.xls` + `config.yaml` | Routing of the 7 study parts identical to the hand-checked workbook (89 of 89 rows, all mapped columns) |
| Factory A, the same data as an Access database (21 tables, one broken calculated column) | one `.accdb` + `config.yaml` | broken column skipped with a message; routing table found among 21 tables; same 89 of 89 rows |
| Factory B, electronics data set (22 CSV files, English headers) | folder + answers for the uncertain columns | 13 sheets, every mapped column identical to the hand-built workbook; one difference: a 24-hour shift ends at `00:00` in the source and `24:00` in the workbook |
| Demo (made up, German headers) | `examples/demo` | interactive run: units from `[min]` and `[h]`, dates to minutes, a column linked from a second file |

---

## Older single-file extractor

`extract.py` is the first version: one routing export, Routing / Products / Machines stubs, configured by `config.yaml`. It still works; `map_schema.py` replaces it.

```bash
python extract.py --input routing_export.xls --config config.yaml --schema schema/Schema_Template_EMPTY.xlsx --output out.xlsx
```

---

## Repository layout

```
erp-to-des-schema/
├── map_schema.py              the tool
├── synonyms.yaml              German / English header synonyms per template column
├── schema/Schema_Template_EMPTY.xlsx
├── config.yaml                example rules file (ERP export with rate codes)
├── examples/
│   ├── demo/                  made-up demo data with German headers
│   ├── factory1_config.yaml   reference copy of config.yaml
│   └── factory2_config.yaml   stub for a CSV routing export
├── extract.py                 older single-file extractor
└── requirements.txt
```
