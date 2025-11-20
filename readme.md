# Alma E-Resource Activation Helper

This is a practical workflow to match your **EBSCO EDS HLM** electronic holdings to the **Alma electronic resources activation form** (macro-enabled Excel workbook).

The goal:

- Let **Python** do all the heavy lifting with big CSVs and name matching.
- Keep **Excel** as the only tool that ever edits the Alma activation form.
- Preserve all **macros, dropdowns, and data validation** in the Alma workbook.

You end up with a checklist that tells you **which rows in the Alma form should have `ACTIVATE? = Y`**, based on your current EDS holdings.

---

## Overview

The workflow has four main phases:

1. Get the latest Alma activation workbook.
2. Export a full EDS holdings CSV.
3. Run three small Python scripts to:
   - Distill EDS title-level data into a package list
   - Extract Alma packages from the macro workbook (read-only)
   - Match EDS packages to Alma packages
4. Use Excel to set `ACTIVATE? = Y` in the Alma workbook using the match reports.

---

## Step 1: Get the latest Alma activation form

You want the latest version of the Alma E-Resource Activation Form so you have up-to-date package names and IDs.

- Go to:  
  https://knowledge.exlibrisgroup.com/Alma/Implementation_and_Migration/Migration_Guides_and_Tutorials/E-Resource_Activations
- Download the current **E-Resource Activation Form** (macro-enabled Excel file, typically `.xlsm`).
- Save it in this project folder as something like:

    E-Resource Activation Form.xlsm

> ⚠️ Treat this file as **sacred**:
> - Never open or edit it in Google Sheets, LibreOffice, or anything other than **Microsoft Excel**.
> - The scripts in this repo only **read** this file; they never write to it.

---

## Step 2: Export holdings from EBSCO EDS

From EBSCOadmin / EDS, export a comprehensive holdings file that includes all selected managed + custom packages.

In the **Download Holdings** interface, use settings like:

- Type: **Standard**
- Format: **CSV**
- Resources: **Managed and Custom**
- Package Content Types: **All**
- Resource Types: **All**
- Visibility: **All Selected Holdings**

Screenshot placeholder — drop your screenshot in the repo and update this path:

    ![EDS Download Holdings screenshot](docs/eds-download-holdings.png)

Save this large CSV into `data/`, for example:

    data/EDS_export.csv

This file is title-level, but it includes **package-level fields** such as:

- `PackageID`
- `PackageName`
- `VendorName`
- `PackageType`
- `PackageAccess`
- `ResourceType`

We will use those fields to infer **which packages you actually subscribe to**.

---

## Project Layout

A suggested directory structure:

    alma_migration/
      E-Resource Activation Form.xlsm       # Alma activation form (macro-enabled, Excel-only)
      requirements.txt
      scripts/
        make_eds_package_list.py
        extract_alma_packages.py
        match_packages.py
      data/
        EDS_export.csv                      # Big EDS holdings CSV from Step 2
      outputs/
        eds_packages_unique.csv
        alma_packages.csv
        matches_exact.csv
        matches_fuzzy.csv
        eds_unmatched.csv
        alma_unmatched.csv
      docs/
        eds-download-holdings.png           # (optional) screenshot for documentation

---

## Environment Setup

Use a Python virtual environment so everything is self-contained.

    python3 -m venv .venv
    source .venv/bin/activate         # Linux/macOS
    # .venv\Scripts\activate          # Windows (PowerShell/cmd)

    pip install --upgrade pip
    pip install -r requirements.txt

`requirements.txt` is very small:

    openpyxl>=3.1

Everything else (`csv`, `sys`, `difflib`, etc.) is from the Python standard library.

---

## Scripts and What They Do

### 1. `make_eds_package_list.py`

**Input:**  
`data/EDS_export.csv` – the large EDS holdings export (title-level).

**Output:**  
`outputs/eds_packages_unique.csv` – one row per distinct package in EDS.

**What it does:**

- Reads the big EDS CSV.
- Uses fields like:
  - `PackageID`
  - `PackageName`
  - `VendorName`
  - `PackageType`
  - `PackageAccess`
  - `ResourceType`
- **Skips** rows where `PackageID` is `0` (e.g., generic “Publisher’s Site” entries that are not true aggregator packages).
- Normalizes package names (lowercase, trims whitespace, simple `&` → `and`).
- Groups by package and counts how many titles are in each package.

Resulting columns (example):

| EDS_PackageID | EDS_PackageName          | EDS_PackageName_norm | VendorName | PackageType | PackageAccess | AnyResourceType | TitleCount |
|---------------|--------------------------|----------------------|-----------|-------------|---------------|-----------------|------------|
| 12345         | Gale General OneFile     | gale general onefile | Gale      | Complete    | Controlled    | Journal         | 10234      |
| 67890         | JSTOR Arts & Sciences IV | jstor arts and...    | JSTOR     | Complete    | Controlled    | Journal         |  1421      |

Run it:

    python scripts/make_eds_package_list.py data/EDS_export.csv outputs/eds_packages_unique.csv

---

### 2. `extract_alma_packages.py`

**Input:**  
`E-Resource Activation Form.xlsm` – the Alma form you downloaded.

**Output:**  
`outputs/alma_packages.csv` – all packages listed in the Alma form’s relevant sheets.

**What it does:**

- Opens the macro-enabled Alma workbook **read-only** using `openpyxl`.
- Reads **three key sheets**:
  - `LICENSED_AGG`
  - `SELECTIVE_PKG`
  - `DATABASE`
- On each sheet, it expects these columns in row 1:
  - `PACKAGE_ID`
  - `PACKAGE_NAME`
  - `SERVICE_TYPE`
- Iterates rows from row 2 downward and stops after hitting a run of empty rows.
- For each non-empty row, writes out:

| SheetName     | ExcelRow | Alma_PACKAGE_ID | Alma_PACKAGE_NAME        | Alma_PACKAGE_NAME_norm | SERVICE_TYPE |
|---------------|----------|-----------------|--------------------------|------------------------|--------------|
| LICENSED_AGG  | 42       | 555555555       | JSTOR Arts & Sciences IV | jstor arts and...      | getFullTxt   |

Notes:

- `ExcelRow` is the 1-based row number in Excel, so you can easily jump to that row later.
- It does **not** modify the `.xlsm` file; it only reads values.

Run it:

    python scripts/extract_alma_packages.py "E-Resource Activation Form.xlsm" outputs/alma_packages.csv

You’ll see a brief summary in the terminal for each sheet, e.g.:

    Sheet LICENSED_AGG: wrote N package rows
    Sheet SELECTIVE_PKG: wrote M package rows
    Sheet DATABASE: wrote K package rows
    Wrote Alma packages to outputs/alma_packages.csv

---

### 3. `match_packages.py`

**Inputs:**

- `outputs/eds_packages_unique.csv` – from Script 1
- `outputs/alma_packages.csv` – from Script 2

**Outputs (in `outputs/`):**

- `matches_exact.csv` – EDS ↔ Alma packages with an exact normalized name match.
- `eds_unmatched.csv` – EDS packages with no Alma match.
- `alma_unmatched.csv` – Alma packages with no EDS match.
- `matches_fuzzy.csv` – optional fuzzy suggestions for EDS packages that didn’t match exactly.

**What it does:**

1. Loads `eds_packages_unique.csv` and indexes by normalized name (`EDS_PackageName_norm`).
2. Loads `alma_packages.csv` and indexes by normalized name (`Alma_PACKAGE_NAME_norm`).
3. **Exact matches:**
   - Where normalized names are identical, records a high-confidence match:

| MatchType  | EDS_PackageName          | EDS_PackageID | VendorName | Alma_PACKAGE_NAME          | Alma_PACKAGE_ID | SheetName    | ExcelRow | SERVICE_TYPE | TitleCount |
|------------|--------------------------|---------------|------------|----------------------------|-----------------|-------------|----------|--------------|-----------|
| exact_name | JSTOR Arts & Sciences IV | 67890         | JSTOR      | JSTOR Arts & Sciences IV   | 555555555       | LICENSED_AGG | 42       | getFullTxt   | 1421      |

4. **Unmatched lists:**
   - `eds_unmatched.csv`: EDS packages with no exact Alma match (useful for spotting missing/renamed packages).
   - `alma_unmatched.csv`: Alma packages with no matching EDS package (potentially dead or new packages you might not want to activate).
5. **Fuzzy suggestions:**
   - For EDS packages with no exact match, calculates a simple similarity between normalized names and suggests the closest Alma package if similarity ≥ 0.75.
   - Output goes into `matches_fuzzy.csv` so you can review “almost matches” by hand.

Run it:

    python scripts/match_packages.py outputs/eds_packages_unique.csv outputs/alma_packages.csv outputs/

You’ll see something like:

    Wrote 400 exact matches to outputs/matches_exact.csv
    Wrote 50 unmatched EDS packages to outputs/eds_unmatched.csv
    Wrote 120 unmatched Alma packages to outputs/alma_unmatched.csv
    Wrote 20 fuzzy matches to outputs/matches_fuzzy.csv

---

## Using the Results in Excel (Manual `Y` Only)

The **key output** for your practical work is:

- `outputs/matches_exact.csv`

Open that in Excel alongside the **Alma activation form**.

For each row in `matches_exact.csv`:

1. Note the `SheetName` (one of `LICENSED_AGG`, `SELECTIVE_PKG`, `DATABASE`).
2. Go to that sheet in **E-Resource Activation Form.xlsm**.
3. Use either:
   - `ExcelRow` to jump directly to the row, or
   - Search for `Alma_PACKAGE_NAME` if that’s more convenient.
4. In that row, set `ACTIVATE?` to `Y` using the dropdown.

Because you only edit the Alma form in **Excel**, and only change the `ACTIVATE?` column:

- All macros remain intact.
- All data validations (dropdowns, allowed values) remain intact.
- You can safely save the workbook and send it back to Ex Libris / your implementation team.

You can also use:

- `matches_fuzzy.csv` to get a list of “maybe this is the same package” items for manual review.
- `eds_unmatched.csv` to see which current EDS packages have no Alma counterpart (or are named very differently).
- `alma_unmatched.csv` to identify Alma packages that you may want to leave turned off.

---

## Safety Notes / Constraints

- **Never** open the Alma workbook in Google Sheets or LibreOffice; that can strip macros and validations.
- The scripts here **never write into `.xlsm`**; they only read from it.
- All activation decisions (`ACTIVATE? = Y`) are made by a human in Excel.

---

## Possible Future Enhancements

- Add a small config file to tweak normalization rules (e.g., special handling for Hein, JSTOR, Brill, etc.).
- Vendor-specific overrides (for known tricky packages) in a separate CSV.
- Reporting/summary script to say “You have X EDS packages, Y matched to Alma, Z unmatched.”

For now, this setup should give you a reproducible, documented way to go from:

> “Huge EDS holdings CSV + complicated Alma activation workbook”

to:

> “Clean checklist and a set of confident `Y` activations in Alma.”
