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

```text
E-Resource Activation Form.xlsm
