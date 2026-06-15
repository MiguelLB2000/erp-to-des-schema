"""
ERP → DES Schema Extractor
Reads an ERP routing export, decodes cycle times, and populates a
DES (Discrete Event Simulation) input schema workbook.

USAGE
  python extract.py \\
      --input  alle_AGs.xls \\
      --config config.yaml \\
      --schema DES_Input_Schema_v04.xlsx \\
      --output Factory1_Schema_v04.xlsx
"""

import argparse
import os
import re
import sys
from collections import defaultdict

import pandas as pd
import openpyxl
import yaml


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(
        description="Extract ERP routing data into a DES input schema workbook.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--input",  required=True, help="Raw ERP export file (.xls or .xlsx)")
    p.add_argument("--config", required=True, help="YAML config file (column mappings, ZBS rules)")
    p.add_argument("--schema", required=True, help="DES schema template workbook (.xlsx)")
    p.add_argument("--output", required=True, help="Output workbook path (.xlsx)")
    return p.parse_args()


# ── Config ────────────────────────────────────────────────────────────────────

def load_config(path):
    with open(path, "r", encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh)
    required = ("columns", "site_filter", "zbs_rules")
    for key in required:
        if key not in cfg:
            print(f"ERROR: config.yaml is missing required section '{key}'")
            sys.exit(1)
    return cfg


# ── Helpers ───────────────────────────────────────────────────────────────────

def to_float(val):
    if pd.isna(val):
        return 0.0
    s = str(val).strip().replace(",", ".").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def decode_cycle_time(kalk_raw, zbs_raw, zbs_cfg):
    """
    Returns (cycle_time_min: float, skip: bool).

    Reads conversion rules from config so different ERP formats with different
    ZBS semantics can be handled without touching code.
    """
    zbs = str(zbs_raw).strip().upper() if not pd.isna(zbs_raw) else ""
    val = to_float(kalk_raw)

    skip_values = {str(v).strip().upper() for v in zbs_cfg.get("skip_values", [])}
    skip_values.add("NAN")  # pandas NaN stringified
    if zbs in skip_values:
        return 0.0, True

    conversions = zbs_cfg.get("conversions", {})
    if zbs in conversions:
        rule = conversions[zbs]
        divisor    = float(rule["divisor"])
        multiplier = float(rule["multiplier"])
        return round((val / divisor) * multiplier, 6), False

    # Unknown ZBS code — keep raw value, don't skip, so the user can review it
    return round(val, 6), False


def make_rg(kst):
    """Kostenstelle → resource-group tag (safe identifier)."""
    return "RG_" + re.sub(r"[^A-Za-z0-9]", "_", str(kst)).upper().strip("_")


def make_mid(fert):
    """Fert.anlage → machine ID tag (safe identifier)."""
    return "M_" + re.sub(r"[^A-Za-z0-9]", "_", str(fert)).upper().strip("_")


def clear_data(ws):
    """Erase everything below row 2 (headers in row 2, metadata in row 1)."""
    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.value = None


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args   = parse_args()
    cfg    = load_config(args.config)
    col    = cfg["columns"]
    zbs_cfg  = cfg["zbs_rules"]
    skip_ops = cfg.get("skip_operation_names", [])
    defaults = cfg.get("defaults", {})
    site_val = cfg["site_filter"]["value"]

    skip_counts = defaultdict(int)   # reason → count

    # ── [1/5] Load ────────────────────────────────────────────────────────────
    print(f"\n[1/5] Loading {args.input} ...")
    if not os.path.exists(args.input):
        print(f"ERROR: {args.input} not found (cwd={os.getcwd()})")
        sys.exit(1)

    try:
        df = pd.read_excel(args.input, engine="xlrd", dtype=str)
    except Exception:
        try:
            df = pd.read_excel(args.input, engine="openpyxl", dtype=str)
        except Exception as exc:
            print(f"ERROR reading {args.input}: {exc}")
            sys.exit(1)

    df.columns = [str(c).strip() for c in df.columns]
    total_loaded = len(df)
    print(f"      {total_loaded:,} rows × {len(df.columns)} columns")

    # Verify required columns exist
    missing = [v for v in col.values() if v not in df.columns]
    if missing:
        print(f"ERROR: columns not found in ERP file: {missing}")
        print(f"       Available columns: {list(df.columns)}")
        sys.exit(1)

    # ── [2/5] Filter ──────────────────────────────────────────────────────────
    print(f"\n[2/5] Filtering rows ...")

    # Site filter
    site_col = col["site"]
    before   = len(df)
    df       = df[df[site_col] == site_val]
    n_site   = before - len(df)
    skip_counts[f"wrong_site (≠{site_val})"] += n_site
    print(f"      {site_col}={site_val}: {len(df):,} rows kept  (removed {n_site:,})")

    # Admin-note operation names (e.g. "Hinweis")
    op_col = col["operation_name"]
    for name in skip_ops:
        mask = df[op_col].str.strip() == name
        n    = int(mask.sum())
        if n:
            df = df[~mask]
            skip_counts[f"operation='{name}'"] += n
            print(f"      After removing '{name}' rows: {len(df):,}")

    # ZBS skip values (cost rows, blank ZBS)
    zbs_col   = col["zbs"]
    skip_zbs  = {str(v).strip().upper() for v in zbs_cfg.get("skip_values", [])}
    skip_zbs.add("NAN")
    skip_mask = df[zbs_col].fillna("").apply(
        lambda z: str(z).strip().upper() in skip_zbs
    )
    n_zbs = int(skip_mask.sum())
    skip_counts["ZBS_skip (cost/blank)"] += n_zbs
    df = df[~skip_mask]
    print(f"      After removing ZBS skip values: {len(df):,}  (dropped {n_zbs:,})")

    # Blank part number or routing step
    part_col = col["part_id"]
    step_col = col["routing_step"]
    before   = len(df)
    df = df[df[part_col].notna() & df[step_col].notna()]
    df = df[df[part_col].str.strip() != ""]
    df = df[df[step_col].str.strip() != ""]
    n_blank = before - len(df)
    if n_blank:
        skip_counts["blank_part_or_step"] += n_blank
    print(f"      After removing blank {part_col}/{step_col}: {len(df):,}")

    # ── [3/5] Deduplicate to latest routing version ────────────────────────────
    plan_col    = col["routing_plan"]
    version_col = col["routing_plan_version"]
    print(f"\n[3/5] Keeping latest {version_col} per {part_col}+{step_col} ...")
    df = df.sort_values(
        [part_col, plan_col, version_col, step_col], na_position="first"
    )
    df = df.drop_duplicates(subset=[part_col, plan_col, step_col], keep="last")
    print(f"      After dedup: {len(df):,} rows")

    print(f"\n      ZBS breakdown in kept rows:")
    print(df[zbs_col].value_counts().to_string())

    print(f"\n      Sample decoded cycle times (first 8 rows):")
    print(f"      {'Teil':>10}  {'AG':>6}  {'kalk.V':>8}  "
          f"{'ZBS':>4}  {'cycle_min':>10}  AG-Bezeichnung")
    for _, r in df.head(8).iterrows():
        ct, _ = decode_cycle_time(r.get(col["calc_rate"]), r.get(zbs_col), zbs_cfg)
        print(f"      {str(r.get(part_col,'')):>10}  "
              f"{str(r.get(step_col,'')):>6}  "
              f"{str(r.get(col['calc_rate'],'')):>8}  "
              f"{str(r.get(zbs_col,'')):>4}  "
              f"{ct:>10.4f}  "
              f"{str(r.get(op_col,''))[:30]}")

    # ── [4/5] Build routing rows ───────────────────────────────────────────────
    print(f"\n[4/5] Building Routing rows ...")

    routing_rows = []
    product_set  = {}            # pid → pname
    kst_map      = {}            # raw kst → RG tag
    machine_set  = defaultdict(set)   # RG tag → {machine IDs}

    for _, row in df.iterrows():
        pid      = str(row.get(col["part_id"],        "")).strip()
        pname    = str(row.get(col["part_name"],      "")).strip()
        step_raw = str(row.get(col["routing_step"],    "")).strip()
        op_name  = str(row.get(col["operation_name"], "")).strip()
        kst      = str(row.get(col["cost_center"],    "")).strip()
        fert     = str(row.get(col["machine_id"],     "")).strip()

        try:
            step_num = int(float(step_raw))
        except ValueError:
            skip_counts["invalid_step_number"] += 1
            continue

        ct, skip = decode_cycle_time(
            row.get(col["calc_rate"]), row.get(zbs_col), zbs_cfg
        )
        if skip:
            skip_counts["ZBS_skip (post-filter)"] += 1
            continue

        rg  = kst_map.setdefault(kst, make_rg(kst))
        mid = make_mid(fert) if fert and fert.lower() not in ("nan", "") else ""
        if mid:
            machine_set[rg].add(mid)

        routing_rows.append({
            "product_id":                  pid,
            "step_number":                 step_num,
            "process_name":                op_name,
            "resource_group":              rg,
            "required_skill":              defaults.get("required_skill", ""),
            "cycle_time_min":              ct,
            "setup_time_min":              defaults.get("setup_time_min", 0.0),
            "batch_overhead_min":          defaults.get("batch_overhead_min", 0.0),
            "batch_size":                  defaults.get("batch_size", 1),
            "yield_pct":                   defaults.get("yield_pct", 100.0),
            "scrap_or_rework":             defaults.get("scrap_or_rework", ""),
            "transport_time_min":          defaults.get("transport_time_min", 0.0),
            "staffing_factor_per_machine": defaults.get("staffing_factor_per_machine", 1.0),
            "operator_attention_factor":   defaults.get("operator_attention_factor", 1.0),
        })
        product_set.setdefault(pid, pname)

    print(f"      Routing rows:               {len(routing_rows):,}")
    print(f"      Unique products:            {len(product_set):,}")
    print(f"      Unique resource groups:     {len(kst_map):,}")

    print(f"\n      KOSTENSTELLE → RESOURCE GROUP  (rename RG_xxx to descriptive names):")
    print(f"      {'Kst.':>10}  →  {'resource_group':<28}  machines (sample)")
    print(f"      {'-'*72}")
    for kst_raw, rg in sorted(kst_map.items(), key=lambda x: x[0]):
        machines = sorted(machine_set.get(rg, []))[:3]
        suffix   = "..." if len(machine_set.get(rg, [])) > 3 else ""
        print(f"      {kst_raw:>10}  →  {rg:<28}  {machines}{suffix}")

    # ── [5/5] Write schema workbook ───────────────────────────────────────────
    print(f"\n[5/5] Writing {args.output} ...")
    if not os.path.exists(args.schema):
        print(f"ERROR: schema template not found: {args.schema}")
        sys.exit(1)

    wb = openpyxl.load_workbook(args.schema)
    fields_populated: set = set()
    fields_todo: set      = set()

    # Routing sheet
    ws_r    = wb["Routing"]
    clear_data(ws_r)
    headers = [c.value for c in ws_r[2]]
    for r_idx, data in enumerate(routing_rows, start=3):
        for c_idx, h in enumerate(headers, start=1):
            if h in data:
                ws_r.cell(row=r_idx, column=c_idx, value=data[h])
                fields_populated.add(h)
    print(f"      Routing: {len(routing_rows):,} rows written")

    # Products sheet (stubs — lot_size / arrival_interval left as TODO)
    ws_p   = wb["Products"]
    clear_data(ws_p)
    prod_h = [c.value for c in ws_p[2]]
    for r_idx, (pid, pname) in enumerate(sorted(product_set.items()), start=3):
        for c_idx, h in enumerate(prod_h, start=1):
            val = {"product_id": pid, "product_name": pname}.get(h, "TODO")
            if val == "TODO":
                fields_todo.add(h)
            ws_p.cell(row=r_idx, column=c_idx, value=val)
    print(f"      Products: {len(product_set):,} stubs")

    # Machines sheet (stubs — MTBF/MTTR left as TODO)
    ws_m   = wb["Machines"]
    clear_data(ws_m)
    mach_h = [c.value for c in ws_m[2]]
    m_row  = 3
    for rg, machines in sorted(machine_set.items()):
        for mid in sorted(machines):
            for c_idx, h in enumerate(mach_h, start=1):
                val = {
                    "machine_id":       mid,
                    "machine_name":     mid,
                    "resource_group":   rg,
                    "station_capacity": 1,
                    "MTBF_hours":       "TODO",
                    "MTTR_min":         "TODO",
                    "cost_per_hour":    0.0,
                }.get(h, "")
                if val == "TODO":
                    fields_todo.add(h)
                ws_m.cell(row=m_row, column=c_idx, value=val)
            m_row += 1
    print(f"      Machines: {m_row - 3} stubs")

    wb.save(args.output)

    # ── Validation summary ────────────────────────────────────────────────────
    total_skipped = sum(skip_counts.values())
    sep = "─" * 54

    print(f"""
{sep}
  VALIDATION SUMMARY
{sep}
  Rows loaded from ERP file     {total_loaded:>10,}
  Rows written to Routing sheet {len(routing_rows):>10,}
  Rows skipped (all reasons)    {total_skipped:>10,}
{sep}
  Skip reasons:""")
    for reason, count in sorted(skip_counts.items(), key=lambda x: -x[1]):
        print(f"    {reason:<36} {count:>8,}")
    print(f"""
  Fields populated              {len(fields_populated):>10}
  Fields still TODO             {len(fields_todo):>10}""")
    if fields_todo:
        for f in sorted(fields_todo):
            print(f"    → {f}")
    print(f"""
  Unique resource groups        {len(kst_map):>10,}
  Unique products               {len(product_set):>10,}
  Machine stubs written         {m_row - 3:>10,}
{sep}
  Output saved to: {args.output}
{sep}

NEXT STEPS
  1. Rename resource groups: replace RG_103, RG_107 etc. with
     descriptive names (e.g. RG_CASTING) using Kostenstellen reference.
  2. MTBF / MTTR: fill in Machines sheet from Stammdaten.xlsx.
  3. Lot sizes: fill lot_size in Products sheet from PB-Familien.xlsx.
  4. Arrival rates: fill arrival_interval_min from Vertriebszahlen.xlsb.
  5. Shifts: populate Shifts sheet from shift schedule data.
  6. setup_time_min: currently 0 — map from a separate source if available.
     (Einlagen column = mold insert codes, NOT setup time.)
""")


if __name__ == "__main__":
    main()
