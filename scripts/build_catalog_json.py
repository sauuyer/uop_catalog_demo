#!/usr/bin/env python3
"""
build_catalog_json.py

Read the two catalog spreadsheets and write catalog.json:

    catalog_sheets/project_metadata.xlsx   one row per project
    catalog_sheets/dataset_metadata.xlsx   one row per NetCDF file

Structure (same shape as the earlier datasets.json: projects -> datasets):

    { "catalog": {...}, "projects": [ { ..., "datasets": [ {...}, ... ] } ] }

Usage:
    python3 build_catalog_json.py --sheets ./catalog_sheets --out ./catalog.json

Requires: openpyxl
"""
import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

SCHEMA_VERSION = "2.0.0"
SEP = "; "

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def read_sheet(path):
    """Return rows as dicts; blank cells and whitespace-only strings -> None."""
    ws = load_workbook(path, read_only=True, data_only=True).active
    rows = ws.iter_rows(values_only=True)
    header = [h for h in next(rows)]
    out = []
    for r in rows:
        if all(v in (None, "") for v in r):
            continue
        d = {}
        for h, v in zip(header, r):
            if isinstance(v, str):
                v = v.strip() or None
            d[h] = v
        out.append(d)
    return out


def split_list(text, sep=SEP):
    if text is None:
        return []
    return [t.strip() for t in str(text).split(sep) if t.strip()]


def unique(items):
    seen, out = set(), []
    for i in items:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def as_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() and isinstance(v, int) else v
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def human_duration(iso):
    """PT900S -> '15 minutes', PT1H -> '1 hour'. Falls back to the raw string."""
    if not iso:
        return None
    m = re.fullmatch(r"P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?",
                     str(iso))
    if not m:
        return str(iso)
    d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    seconds = d * 86400 + h * 3600 + mi * 60 + s
    if seconds == 0:
        return str(iso)
    for unit, size in (("day", 86400), ("hour", 3600), ("minute", 60)):
        if seconds % size == 0:
            n = seconds // size
            return f"{n} {unit}{'' if n == 1 else 's'}"
    if seconds >= 60 and seconds % 30 == 0:  # e.g. 450 s -> 7.5 minutes
        return f"{seconds / 60:g} minutes"
    return str(iso)


def acronym_from(name):
    """'Coastal Mixing and Optics Experiment (CMO)' -> 'CMO'."""
    m = re.search(r"\(([A-Za-z0-9\-]+)\)\s*$", name or "")
    return m.group(1) if m else None


def location_block(lat_min, lat_max, lon_min, lon_max):
    if None in (lat_min, lat_max, lon_min, lon_max):
        return None
    block = {"bbox": {"lat_min": lat_min, "lat_max": lat_max,
                      "lon_min": lon_min, "lon_max": lon_max}}
    if lat_min == lat_max and lon_min == lon_max:
        block["latitude"], block["longitude"] = lat_min, lon_min
    return block


# --------------------------------------------------------------------------
# Record builders
# --------------------------------------------------------------------------


def build_dataset(r, project):
    codes = split_list(r["variable_codes"])
    long_names = split_list(r["variables"])
    # standard_names / units can contain empty slots, so keep positions
    std = [s.strip() for s in str(r["variable_standard_names"] or "").split(";")]
    units = [u.strip() for u in str(r["variable_units"] or "").split(";")]
    std += [""] * (len(codes) - len(std))
    units += [""] * (len(codes) - len(units))
    details = [{
        "code": c,
        "long_name": long_names[i] if i < len(long_names) else c,
        "standard_name": std[i] or None,
        "units": units[i] or None,
    } for i, c in enumerate(codes)]

    creator = None
    if r["creator_name"]:
        creator = {"name": r["creator_name"], "email": r["creator_email"],
                   "url": r["creator_url"],
                   "institution": r["creator_institution"]}
    publisher = None
    if r["publisher_name"]:
        publisher = {"name": r["publisher_name"], "url": r["publisher_url"],
                     "institution": r["publisher_institution"]}

    depth = None
    if r["depth_min"] is not None or r["depth_levels"]:
        depth = {"min": as_num(r["depth_min"]), "max": as_num(r["depth_max"]),
                 "units": "m", "positive": "down",
                 "levels": [as_num(x) for x in split_list(r["depth_levels"])],
                 "n_levels": as_num(r["n_depth_levels"])}

    return {
        "id": r["dataset_id"],
        "naming_authority": r["naming_authority"],
        "title": r["title"],
        "summary": r["summary"],
        "status": project["status"],
        "keywords": split_list(r["keywords"], sep=SEP),
        "keywords_vocabulary": r["keywords_vocabulary"],
        "region": project["region"],
        "time_start": r["time_start"],
        "time_end": r["time_end"],
        "granularity": {
            "temporal_resolution": human_duration(r["temporal_resolution"]),
            "temporal_resolution_iso": r["temporal_resolution"],
            "n_samples": as_num(r["n_time_samples"]),
            "time_coverage_duration": r["time_duration"],
        },
        "location": location_block(as_num(r["lat_min"]), as_num(r["lat_max"]),
                                   as_num(r["lon_min"]), as_num(r["lon_max"])),
        "depth": depth,
        "variables": unique(long_names),
        "variable_details": details,
        "creator": creator,
        "publisher": publisher,
        "institution": r["institution"],
        "license": r["license"],
        "citation": r["citation"],
        "acknowledgement": r["acknowledgement"],
        "conventions": r["conventions"],
        "standard_name_vocabulary": r["standard_name_vocabulary"],
        "record_updated": r["record_updated"],
        "file": {
            "name": r["file_name"],
            "relative_path": r["relative_path"],
            "format": "NetCDF",
            "size_bytes": as_num(r["file_size_bytes"]),
            "sha256": r["checksum_sha256"],
        },
        "links": [],
    }


def build_project(p, dataset_rows):
    project = {
        "id": p["project_id"],
        "acronym": acronym_from(p["project_name"]),
        "title": p["project_name"],
        "status": (p["project_status"] or "").lower() or None,
        "summary": p["project_summary"],
        "region": p["region"],
    }
    datasets = [build_dataset(r, project) for r in dataset_rows]

    points = unique((d["location"]["latitude"], d["location"]["longitude"])
                    for d in datasets
                    if d["location"] and "latitude" in d["location"])
    loc = location_block(as_num(p["lat_min"]), as_num(p["lat_max"]),
                         as_num(p["lon_min"]), as_num(p["lon_max"]))
    if loc is not None:
        loc["points"] = [{"latitude": a, "longitude": b} for a, b in points]
        if len(points) == 1:
            loc["latitude"], loc["longitude"] = points[0]
    project.update({
        "location": loc,
        "time_start": p["time_start"],
        "time_end": p["time_end"],
        "keywords": split_list(p["keywords"]),
        "creators": split_list(p["creators"]),
        "institutions": split_list(p["institutions"]),
        "publishers": split_list(p["publishers"]),
        "licenses": split_list(p["licenses"]),
        "n_datasets": len(datasets),
        "links": [],
        "datasets": datasets,
    })
    return project


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--sheets", default="catalog_sheets")
    ap.add_argument("--out", default="catalog.json")
    args = ap.parse_args()

    sheets = Path(args.sheets)
    projects = read_sheet(sheets / "project_metadata.xlsx")
    datasets = read_sheet(sheets / "dataset_metadata.xlsx")

    problems = []
    proj_ids = [p["project_id"] for p in projects]
    ds_ids = [d["dataset_id"] for d in datasets]
    for dup in {i for i in ds_ids if ds_ids.count(i) > 1}:
        problems.append(f"duplicate dataset_id: {dup}")
    for d in datasets:
        if d["project_id"] not in proj_ids:
            problems.append(f"dataset {d['dataset_id']} has unknown "
                            f"project_id {d['project_id']}")
    for p in projects:
        for col in ("project_summary", "project_status", "region"):
            if not p[col]:
                problems.append(f"project {p['project_id']}: {col} is empty")
        n = sum(1 for d in datasets if d["project_id"] == p["project_id"])
        if p["n_datasets"] is not None and int(p["n_datasets"]) != n:
            problems.append(f"project {p['project_id']}: sheet says "
                            f"{p['n_datasets']} datasets, found {n}")
    fatal = [x for x in problems if "unknown" in x or "duplicate" in x]
    if fatal:
        sys.exit("ERROR:\n  " + "\n  ".join(fatal))

    out_projects = [
        build_project(p, [d for d in datasets
                          if d["project_id"] == p["project_id"]])
        for p in projects
    ]

    catalog = {
        "catalog": {
            "id": "uop_catalog",
            "title": "Upper Ocean Processes (UOP) Data Catalog",
            "description": ("Catalog of UOP projects and datasets, built from "
                            "the metadata inside the processed NetCDF files."),
            "generated_utc": dt.datetime.now(dt.timezone.utc).strftime(
                "%Y-%m-%dT%H:%M:%SZ"),
            "schema_version": SCHEMA_VERSION,
            "n_projects": len(out_projects),
            "n_datasets": sum(p["n_datasets"] for p in out_projects),
        },
        "projects": out_projects,
    }

    Path(args.out).write_text(json.dumps(catalog, indent=2, ensure_ascii=False)
                              + "\n", encoding="utf-8")
    print(f"wrote {args.out}: {catalog['catalog']['n_projects']} projects, "
          f"{catalog['catalog']['n_datasets']} datasets")
    for x in problems:
        print("  warning:", x)


if __name__ == "__main__":
    main()
