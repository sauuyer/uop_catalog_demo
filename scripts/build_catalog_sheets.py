#!/usr/bin/env python3
"""
build_catalog_sheets.py

Walk processed_uop_netcdfs/<project>/netCDF/*.nc and write two workbooks that
feed the UOP data catalog JSON:

    dataset_metadata.xlsx   one row per NetCDF file
    project_metadata.xlsx   one row per project folder

Everything that can be read from the files is read from the files. Fields the
NetCDFs do not carry (platform, instrument, region, project summary, project
status) get yellow "manual" columns. Re-running the script keeps whatever you
typed in those columns (matched on dataset_id / project_id).

Usage:
    python3 build_catalog_sheets.py --root ./processed_uop_netcdfs --out ./catalog_sheets

Requires: netCDF4, numpy, openpyxl
"""
import argparse
import datetime as dt
import hashlib
import re
import sys
from collections import Counter
from pathlib import Path

import netCDF4
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Column definitions
# --------------------------------------------------------------------------

# Columns the NetCDFs cannot supply. Filled by hand, preserved across re-runs.
DATASET_MANUAL = ["platform"]
PROJECT_MANUAL = ["region", "project_summary", "project_status"]

DATASET_COLUMNS = [
    # identity
    "dataset_id", "file_id", "naming_authority", "project_id", "project_folder",
    "project_name", "file_name", "relative_path",
    # descriptive
    "title", "summary", "keywords", "keywords_vocabulary",
    # time
    "time_start", "time_end", "time_duration", "temporal_resolution",
    "n_time_samples",
    # space
    "lat_min", "lat_max", "lon_min", "lon_max",
    "depth_min", "depth_max", "depth_levels", "n_depth_levels",
    # people / rights
    "creator_name", "creator_email", "creator_url", "creator_institution",
    "publisher_name", "publisher_url", "publisher_institution", "institution",
    "license", "citation", "acknowledgement",
    # standards / provenance
    "conventions", "standard_name_vocabulary", "record_updated",
    # variables (semicolon-separated, aligned by position)
    "variables", "variable_codes", "variable_standard_names", "variable_units",
    # file facts
    "file_size_bytes", "checksum_sha256",
    # not in the NetCDFs: fill in by hand
    *DATASET_MANUAL,
    # QA
    "validation_warnings",
]

PROJECT_COLUMNS = [
    "project_id", "project_folder", "project_name",
    "time_start", "time_end",
    "lat_min", "lat_max", "lon_min", "lon_max", "n_locations",
    "n_datasets", "dataset_ids",
    "keywords", "creators", "institutions", "publishers", "licenses",
    "acknowledgements", "citations",
    # not in the NetCDFs: fill in by hand
    *PROJECT_MANUAL,
    "validation_warnings",
]

# Global attributes every file should carry. Missing ones are flagged.
REQUIRED_GLOBALS = [
    "title", "summary", "keywords", "project", "id", "naming_authority",
    "creator_name", "institution", "publisher_name", "license", "citation",
    "time_coverage_start", "time_coverage_end", "time_coverage_resolution",
    "geospatial_lat_min", "geospatial_lat_max",
    "geospatial_lon_min", "geospatial_lon_max",
]

COORD_VARS = {"TIME", "DEPTH", "LATITUDE", "LONGITUDE"}
PLACEHOLDER_VALUES = {"", "not available", "n/a", "na", "none", "unknown"}

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def slug(text):
    return re.sub(r"[^a-z0-9]+", "_", str(text).lower()).strip("_")


def get_attr(obj, name, default=None):
    if name in obj.ncattrs():
        val = obj.getncattr(name)
        if isinstance(val, bytes):
            val = val.decode("utf-8", "replace")
        if isinstance(val, np.generic):
            val = val.item()
        return val
    return default


def to_float(val):
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def is_placeholder(val):
    return val is None or str(val).strip().lower() in PLACEHOLDER_VALUES


def fmt_num(x, nd=6):
    if x is None:
        return None
    return round(float(x), nd)


def epoch_to_iso(seconds):
    return dt.datetime.fromtimestamp(float(seconds), dt.timezone.utc).strftime(
        "%Y-%m-%dT%H:%M:%SZ")


def parse_iso(s):
    if not s:
        return None
    try:
        return dt.datetime.strptime(str(s), "%Y-%m-%dT%H:%M:%SZ").replace(
            tzinfo=dt.timezone.utc)
    except ValueError:
        return None


def sha256_of(path, chunk=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def split_keywords(text):
    """Keywords are stored as one comma-separated string."""
    if not text:
        return []
    return [k.strip() for k in str(text).split(",") if k.strip()]


def unique_in_order(items):
    seen, out = set(), []
    for it in items:
        if it and it not in seen:
            seen.add(it)
            out.append(it)
    return out


def join(items, sep="; "):
    return sep.join(str(i) for i in items if i not in (None, ""))


def time_seconds(ds):
    """Valid TIME values as epoch seconds, or None if units are unexpected."""
    if "TIME" not in ds.variables:
        return None
    var = ds.variables["TIME"]
    units = str(get_attr(var, "units", ""))
    if not units.startswith("seconds since 1970-01-01"):
        return None
    var.set_auto_mask(True)
    arr = np.ma.masked_invalid(var[:]).compressed()
    return arr if arr.size else None


def coord_values(ds, name):
    if name not in ds.variables:
        return None
    var = ds.variables[name]
    var.set_auto_mask(True)
    arr = np.ma.masked_invalid(var[:]).compressed()
    return arr if arr.size else None


# --------------------------------------------------------------------------
# Per-file extraction
# --------------------------------------------------------------------------


def extract_file(nc_path, root):
    warnings = []
    project_folder = nc_path.parent.parent.name
    rel = nc_path.relative_to(root.parent) if root.parent in nc_path.parents \
        else nc_path
    row = {c: None for c in DATASET_COLUMNS}

    with netCDF4.Dataset(nc_path) as ds:
        g = lambda name: get_attr(ds, name)  # noqa: E731

        # --- identity ------------------------------------------------------
        raw_id = g("id")
        row["file_id"] = raw_id
        base_id = re.sub(r"_nc$", "", str(raw_id)) if raw_id else nc_path.stem
        row["dataset_id"] = base_id
        row["naming_authority"] = g("naming_authority")
        row["project_folder"] = project_folder
        row["project_id"] = slug(project_folder)
        row["project_name"] = g("project")
        row["file_name"] = nc_path.name
        row["relative_path"] = str(rel).replace("\\", "/")

        # --- descriptive ---------------------------------------------------
        row["title"] = g("title")
        row["summary"] = g("summary")
        row["keywords"] = join(split_keywords(g("keywords")))
        row["keywords_vocabulary"] = g("keywords_vocabulary")

        # --- time ----------------------------------------------------------
        row["time_start"] = g("time_coverage_start")
        row["time_end"] = g("time_coverage_end")
        row["time_duration"] = g("time_coverage_duration")
        row["temporal_resolution"] = g("time_coverage_resolution")
        secs = time_seconds(ds)
        row["n_time_samples"] = int(len(ds.dimensions["TIME"])) \
            if "TIME" in ds.dimensions else None
        if secs is not None:
            actual_start, actual_end = epoch_to_iso(secs.min()), \
                epoch_to_iso(secs.max())
            for label, attr_val, actual in (
                    ("time_coverage_start", row["time_start"], actual_start),
                    ("time_coverage_end", row["time_end"], actual_end)):
                a, b = parse_iso(attr_val), parse_iso(actual)
                if a and b and abs((a - b).total_seconds()) > 3600:
                    warnings.append(
                        f"{label} attr {attr_val} differs from TIME data {actual}")
            if not row["time_start"]:
                row["time_start"] = actual_start
            if not row["time_end"]:
                row["time_end"] = actual_end
        else:
            warnings.append("TIME variable missing or units not seconds since 1970")

        # --- space ---------------------------------------------------------
        lat, lon = coord_values(ds, "LATITUDE"), coord_values(ds, "LONGITUDE")
        for key, attr, vals, fn in (
                ("lat_min", "geospatial_lat_min", lat, np.min),
                ("lat_max", "geospatial_lat_max", lat, np.max),
                ("lon_min", "geospatial_lon_min", lon, np.min),
                ("lon_max", "geospatial_lon_max", lon, np.max)):
            a = to_float(g(attr))
            v = float(fn(vals)) if vals is not None else None
            if a is not None and v is not None and abs(a - v) > 0.01:
                warnings.append(f"{attr} {a:.4f} differs from variable {v:.4f}")
            row[key] = fmt_num(a if a is not None else v)

        depth = coord_values(ds, "DEPTH")
        if depth is not None:
            levels = sorted(set(round(float(d), 2) for d in depth))
            row["depth_levels"] = join(
                [int(d) if float(d).is_integer() else d for d in levels])
            row["n_depth_levels"] = len(levels)
            dmin, dmax = levels[0], levels[-1]
        else:
            dmin = dmax = None
        vmin, vmax = to_float(g("geospatial_vertical_min")), \
            to_float(g("geospatial_vertical_max"))
        row["depth_min"] = fmt_num(vmin if vmin is not None else dmin, 3)
        row["depth_max"] = fmt_num(vmax if vmax is not None else dmax, 3)
        if vmin is not None and dmin is not None and abs(vmin - dmin) > 0.01:
            warnings.append(
                f"geospatial_vertical_min {vmin} differs from DEPTH {dmin}")
        if vmax is not None and dmax is not None and abs(vmax - dmax) > 0.01:
            warnings.append(
                f"geospatial_vertical_max {vmax} differs from DEPTH {dmax}")

        # --- people / rights ----------------------------------------------
        for col in ("creator_name", "creator_email", "creator_url",
                    "creator_institution", "publisher_name", "publisher_url",
                    "publisher_institution", "institution", "license",
                    "citation", "acknowledgement"):
            row[col] = g(col)

        # --- standards / provenance --------------------------------------
        row["conventions"] = g("Conventions")
        row["standard_name_vocabulary"] = g("standard_name_vocabulary")
        row["record_updated"] = g("date_created")

        # --- variables -----------------------------------------------------
        codes, std, longn, units = [], [], [], []
        for name, var in ds.variables.items():
            if name in COORD_VARS:
                continue
            sn, ln, un = (get_attr(var, "standard_name"),
                          get_attr(var, "long_name"), get_attr(var, "units"))
            missing = [lab for lab, v in (("standard_name", sn),
                                          ("long_name", ln), ("units", un))
                       if is_placeholder(v)]
            if missing:
                warnings.append(f"variable {name}: missing/placeholder "
                                + ", ".join(missing))
            codes.append(name)
            std.append("" if is_placeholder(sn) else sn)
            longn.append(name if is_placeholder(ln) else ln)
            units.append("" if is_placeholder(un) else un)
        row["variable_codes"] = join(codes)
        row["variables"] = join(longn)
        row["variable_standard_names"] = "; ".join(std)
        row["variable_units"] = "; ".join(units)

        # --- required globals ---------------------------------------------
        for attr in REQUIRED_GLOBALS:
            if is_placeholder(g(attr)):
                warnings.append(f"missing global attribute: {attr}")

    # --- file facts --------------------------------------------------------
    row["file_size_bytes"] = nc_path.stat().st_size
    row["checksum_sha256"] = sha256_of(nc_path)
    row["validation_warnings"] = " | ".join(warnings)
    return row


# --------------------------------------------------------------------------
# Project roll-up
# --------------------------------------------------------------------------


def rollup_projects(dataset_rows):
    by_proj = {}
    for r in dataset_rows:
        by_proj.setdefault(r["project_folder"], []).append(r)

    out = []
    for folder, rows in sorted(by_proj.items()):
        warnings = []
        names = [r["project_name"] for r in rows if r["project_name"]]
        name_counts = Counter(names)
        if len(name_counts) > 1:
            warnings.append("project attribute differs between files: "
                            + "; ".join(f"{n!r} x{c}" for n, c in
                                        name_counts.items()))
        name = name_counts.most_common(1)[0][0] if names else folder

        def span(col, fn):
            vals = [r[col] for r in rows if r[col] is not None]
            return fn(vals) if vals else None

        points = {(r["lat_min"], r["lon_min"]) for r in rows
                  if r["lat_min"] is not None and r["lon_min"] is not None}

        def union(col):
            return unique_in_order([r[col] for r in rows])

        kw = unique_in_order(k for r in rows
                             for k in (r["keywords"] or "").split("; "))
        out.append({
            "project_id": slug(folder),
            "project_folder": folder,
            "project_name": name,
            "time_start": span("time_start", min),
            "time_end": span("time_end", max),
            "lat_min": span("lat_min", min), "lat_max": span("lat_max", max),
            "lon_min": span("lon_min", min), "lon_max": span("lon_max", max),
            "n_locations": len(points),
            "n_datasets": len(rows),
            "dataset_ids": join(r["dataset_id"] for r in rows),
            "keywords": join(kw),
            "creators": join(union("creator_name")),
            "institutions": join(union("institution")),
            "publishers": join(union("publisher_name")),
            "licenses": join(union("license")),
            "acknowledgements": join(union("acknowledgement")),
            "citations": join(union("citation"), sep=" || "),
            "project_summary": None, "project_status": None, "region": None,
            "validation_warnings": " | ".join(warnings),
        })
    return out


# --------------------------------------------------------------------------
# Workbook I/O
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
MANUAL_FILL = PatternFill("solid", fgColor="FFF2CC")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
WIDE = {"summary": 60, "title": 50, "citation": 50, "citations": 50,
        "keywords": 40, "validation_warnings": 50, "project_summary": 60,
        "variables": 40, "variable_codes": 24, "dataset_ids": 40,
        "relative_path": 40, "checksum_sha256": 20, "acknowledgement": 36,
        "acknowledgements": 36}


def load_manual(path, key_col, manual_cols):
    """Read previously entered manual values, keyed by key_col."""
    saved = {}
    if not path.exists():
        return saved
    ws = load_workbook(path, read_only=True).active
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header or key_col not in header:
        return saved
    idx = {h: i for i, h in enumerate(header)}
    for r in rows:
        key = r[idx[key_col]]
        if key is None:
            continue
        saved[key] = {c: r[idx[c]] for c in manual_cols if c in idx
                      and r[idx[c]] not in (None, "")}
    return saved


def write_workbook(path, columns, rows, manual_cols, sheet_name):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(columns)
    for r in rows:
        ws.append([r.get(c) for c in columns])

    for ci, col in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=ci)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = MANUAL_FILL if col in manual_cols else HEADER_FILL
        if col in manual_cols:
            cell.font = Font(bold=True, color="000000")
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = WIDE.get(
            col, max(14, min(28, len(col) + 4)))
        for ri in range(2, len(rows) + 2):
            c = ws.cell(row=ri, column=ci)
            c.alignment = Alignment(vertical="top", wrap_text=col in WIDE)
            if col in manual_cols:
                c.fill = MANUAL_FILL
            elif col == "validation_warnings" and c.value:
                c.fill = WARN_FILL
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--root", default="processed_uop_netcdfs",
                    help="folder holding <project>/netCDF/*.nc")
    ap.add_argument("--out", default="catalog_sheets",
                    help="output folder for the two .xlsx files")
    ap.add_argument("--no-checksum", action="store_true",
                    help="skip SHA-256 (faster for quick checks)")
    args = ap.parse_args()

    root = Path(args.root).resolve()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    files = sorted(root.glob("*/netCDF/*.nc"))
    if not files:
        sys.exit(f"No files found at {root}/*/netCDF/*.nc")

    ds_path, pr_path = out / "dataset_metadata.xlsx", out / "project_metadata.xlsx"
    saved_ds = load_manual(ds_path, "dataset_id", DATASET_MANUAL)
    saved_pr = load_manual(pr_path, "project_id", PROJECT_MANUAL)

    dataset_rows = []
    for f in files:
        print(f"reading {f.relative_to(root)}")
        row = extract_file(f, root)
        if args.no_checksum:
            row["checksum_sha256"] = None
        row.update(saved_ds.get(row["dataset_id"], {}))
        dataset_rows.append(row)

    ids = Counter(r["dataset_id"] for r in dataset_rows)
    for r in dataset_rows:
        if ids[r["dataset_id"]] > 1:
            r["validation_warnings"] = (r["validation_warnings"] + " | "
                                        if r["validation_warnings"] else "") \
                + "duplicate dataset_id"

    project_rows = rollup_projects(dataset_rows)
    for r in project_rows:
        r.update(saved_pr.get(r["project_id"], {}))

    write_workbook(ds_path, DATASET_COLUMNS, dataset_rows, DATASET_MANUAL,
                   "datasets")
    write_workbook(pr_path, PROJECT_COLUMNS, project_rows, PROJECT_MANUAL,
                   "projects")

    n_warn = sum(1 for r in dataset_rows if r["validation_warnings"])
    print(f"\n{len(dataset_rows)} files, {len(project_rows)} projects")
    print(f"{n_warn} file(s) with validation warnings")
    print(f"wrote {ds_path}\nwrote {pr_path}")


if __name__ == "__main__":
    main()
