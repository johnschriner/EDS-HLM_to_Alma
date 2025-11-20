#!/usr/bin/env python3
"""
make_eds_package_list.py

Read a very large EDS holdings CSV (title-level) and produce a
de-duplicated list of packages with counts.

Usage:
    python make_eds_package_list.py path/to/EDS_export.csv outputs/eds_packages_unique.csv
"""

import csv
import sys
from collections import defaultdict

def normalize(s: str) -> str:
    if s is None:
        return ""
    s = s.strip().lower()
    # simple normalization for & vs "and"
    s = s.replace("&", "and")
    while "  " in s:
        s = s.replace("  ", " ")
    return s

def main(in_path: str, out_path: str) -> None:
    # We only need a few columns; names based on your sample
    key_cols = [
        "PackageID",
        "PackageName",
        "VendorName",
        "PackageType",
        "PackageAccess",
        "ResourceType",
    ]

    # Use dict keyed by all of the above plus normalized name
    packages = {}
    counts = defaultdict(int)

    with open(in_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        missing = [c for c in key_cols if c not in reader.fieldnames]
        if missing:
            raise SystemExit(f"ERROR: Missing expected columns in EDS CSV: {missing}")

        for row in reader:
            pkg_id = row["PackageID"].strip() if row["PackageID"] else ""
            # Skip non-packages (PackageID == 0) – mostly "Publisher's Site"
            if pkg_id in ("", "0"):
                continue

            pkg_name = (row["PackageName"] or "").strip()
            vendor = (row["VendorName"] or "").strip()
            pkg_type = (row["PackageType"] or "").strip()
            pkg_access = (row["PackageAccess"] or "").strip()
            rtype = (row["ResourceType"] or "").strip()

            norm_name = normalize(pkg_name)

            key = (pkg_id, norm_name, vendor, pkg_type, pkg_access)
            counts[key] += 1
            if key not in packages:
                packages[key] = {
                    "EDS_PackageID": pkg_id,
                    "EDS_PackageName": pkg_name,
                    "EDS_PackageName_norm": norm_name,
                    "VendorName": vendor,
                    "PackageType": pkg_type,
                    "PackageAccess": pkg_access,
                    "AnyResourceType": rtype,  # just the first we see
                }

    # Write out de-duplicated packages
    fieldnames = [
        "EDS_PackageID",
        "EDS_PackageName",
        "EDS_PackageName_norm",
        "VendorName",
        "PackageType",
        "PackageAccess",
        "AnyResourceType",
        "TitleCount",
    ]

    with open(out_path, "w", newline="", encoding="utf-8") as out_f:
        writer = csv.DictWriter(out_f, fieldnames=fieldnames)
        writer.writeheader()
        for key, info in packages.items():
            info = info.copy()
            info["TitleCount"] = counts[key]
            writer.writerow(info)

    print(f"Wrote {len(packages)} unique packages to {out_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python make_eds_package_list.py EDS_export.csv outputs/eds_packages_unique.csv")
        raise SystemExit(1)
    main(sys.argv[1], sys.argv[2])
