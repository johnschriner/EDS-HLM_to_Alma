#!/usr/bin/env python3
"""
match_packages.py

Match EDS unique packages to Alma packages (by normalized name).

Usage:
    python match_packages.py outputs/eds_packages_unique.csv outputs/alma_packages.csv outputs/
"""

import sys
import csv

def main(eds_path: str, alma_path: str, out_dir: str) -> None:
    # --- Load EDS packages ---
    eds_packages = {}
    with open(eds_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            norm = row.get("EDS_PackageName_norm", "").strip()
            if not norm:
                continue
            # If there are duplicates on norm, keep the first one (that’s fine for our use)
            if norm not in eds_packages:
                eds_packages[norm] = row

    # --- Load Alma packages ---
    alma_packages = {}
    alma_by_norm = {}
    with open(alma_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            norm = row.get("Alma_PACKAGE_NAME_norm", "").strip()
            if not norm:
                continue
            idx = len(alma_packages)
            alma_packages[idx] = row
            alma_by_norm.setdefault(norm, []).append(idx)

    # --- Exact matches ---
    matches_exact = []
    eds_norms_seen = set()
    alma_idxs_seen = set()

    for norm, eds_row in eds_packages.items():
        if norm in alma_by_norm:
            for idx in alma_by_norm[norm]:
                alma_row = alma_packages[idx]
                match = {
                    "MatchType": "exact_name",
                    "EDS_PackageName": eds_row.get("EDS_PackageName", ""),
                    "EDS_PackageID": eds_row.get("EDS_PackageID", ""),
                    "VendorName": eds_row.get("VendorName", ""),
                    "Alma_PACKAGE_NAME": alma_row.get("Alma_PACKAGE_NAME", ""),
                    "Alma_PACKAGE_ID": alma_row.get("Alma_PACKAGE_ID", ""),
                    "SheetName": alma_row.get("SheetName", ""),
                    "ExcelRow": alma_row.get("ExcelRow", ""),
                    "SERVICE_TYPE": alma_row.get("SERVICE_TYPE", ""),
                    "TitleCount": eds_row.get("TitleCount", ""),
                }
                matches_exact.append(match)
                eds_norms_seen.add(norm)
                alma_idxs_seen.add(idx)

    # --- Unmatched sets ---
    unmatched_eds = [
        row for norm, row in eds_packages.items() if norm not in eds_norms_seen
    ]
    unmatched_alma = [
        row for idx, row in alma_packages.items() if idx not in alma_idxs_seen
    ]

    # --- Prepare output paths ---
    out_dir = out_dir.rstrip("/\\")
    exact_out = f"{out_dir}/matches_exact.csv"
    eds_unmatched_out = f"{out_dir}/eds_unmatched.csv"
    alma_unmatched_out = f"{out_dir}/alma_unmatched.csv"

    # --- Write exact matches ---
    if matches_exact:
        fieldnames = list(matches_exact[0].keys())
        with open(exact_out, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(matches_exact)
    print(f"Wrote {len(matches_exact)} exact matches to {exact_out}")

    # --- Write unmatched EDS ---
    if unmatched_eds:
        fieldnames = list(unmatched_eds[0].keys())
        with open(eds_unmatched_out, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(unmatched_eds)
    print(f"Wrote {len(unmatched_eds)} unmatched EDS packages to {eds_unmatched_out}")

    # --- Write unmatched Alma ---
    if unmatched_alma:
        fieldnames = list(unmatched_alma[0].keys())
        with open(alma_unmatched_out, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(unmatched_alma)
    print(f"Wrote {len(unmatched_alma)} unmatched Alma packages to {alma_unmatched_out}")

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python match_packages.py outputs/eds_packages_unique.csv outputs/alma_packages.csv outputs/")
        raise SystemExit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])
