"""
parser.py — SimaPro LCIA CSV parser.

SimaPro exports are not standard CSVs. A metadata header block (project name,
date, method, etc.) precedes the actual data table. This module locates the
data table, auto-detects the field delimiter, and returns a clean DataFrame.

SimaPro CSV structure (simplified):
    SimaPro 9.x.x
    Project: My Project
    Date: 01/01/2024
    ...                          <-- metadata rows (ignored)
    Impact category;Unit;S1;S2  <-- first cell == "Impact category" → data start
    Global warming;kg CO2 eq;100;200
    ...
    End                          <-- SimaPro section terminator

Assumptions about the SimaPro export format are flagged with # ASSUMPTION comments.
"""

import csv
import io
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd


# ASSUMPTION: SimaPro uses one of these labels as the first column of the data
# table depending on the indicator type selected at export time:
#   "Impact category"  — midpoint / characterisation exports
#   "Damage category"  — endpoint / damage assessment exports
# Add any additional labels your SimaPro version uses to this set.
_IMPACT_CATEGORY_HEADERS = {"impact category", "damage category"}


def _detect_delimiter(raw_text: str) -> str:
    """
    Auto-detect whether the CSV uses comma or semicolon as field delimiter.

    SimaPro uses commas in English-locale installations and semicolons in
    European locales (where a comma is the decimal separator).

    Strategy: count ';' vs ',' in the first 30 lines. The more frequent
    character is the delimiter.

    Args:
        raw_text: Full file contents as a string.

    Returns:
        Either ',' or ';'.
    """
    head = "\n".join(raw_text.splitlines()[:30])
    counts = {
        "\t": head.count("\t"),
        ";":  head.count(";"),
        ",":  head.count(","),
    }
    return max(counts, key=counts.get)


def _find_data_header_row(lines: List[str], delimiter: str) -> Optional[int]:
    """
    Scan lines to find the zero-based index of the data table header row.

    SimaPro places a row whose first cell equals "Impact category"
    (case-insensitive) immediately before the data rows. All rows above it
    are metadata and are ignored.

    Args:
        lines: All lines of the file as strings.
        delimiter: The detected field delimiter.

    Returns:
        Index of the header row, or None if not found.
    """
    for i, line in enumerate(lines):
        first_cell = line.split(delimiter)[0].strip().strip('"').lower()
        if first_cell in _IMPACT_CATEGORY_HEADERS:
            return i
    return None


def parse_simapro_csv(file_path: Path) -> pd.DataFrame:
    """
    Parse a SimaPro LCIA CSV export into a clean pandas DataFrame.

    The returned DataFrame has the following columns:
        - impact_category (str): Name of the LCIA impact category.
        - unit (str): Unit of measurement (e.g. 'kg CO2 eq').
        - <scenario_name> (float): One column per product system / scenario in
          the export, containing the characterisation result for that category.

    Args:
        file_path: Path to the SimaPro CSV export file.

    Returns:
        A pandas DataFrame with one row per impact category.

    Raises:
        FileNotFoundError: If the CSV file does not exist at the given path.
        ValueError: If the file cannot be parsed as a valid SimaPro LCIA CSV.
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(
            f"CSV file not found at: {file_path}\n"
            "Check that the path entered in the Settings sheet is correct."
        )

    # Try UTF-8 with BOM first (utf-8-sig strips the BOM automatically).
    # Fall back to latin-1, which some older SimaPro versions output.
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            raw_text = file_path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(
            f"Could not decode the file: {file_path}\n"
            "Try re-exporting from SimaPro with UTF-8 encoding."
        )

    delimiter = _detect_delimiter(raw_text)
    # str.splitlines() correctly handles both LF and CRLF line endings.
    lines = raw_text.splitlines()

    header_row_index = _find_data_header_row(lines, delimiter)

    if header_row_index is None:
        raise ValueError(
            "Could not locate the LCIA data table in the file.\n"
            "Expected a row whose first cell is 'Impact category' or 'Damage category'.\n"
            "If your SimaPro version uses a different label, add it to "
            "_IMPACT_CATEGORY_HEADERS in lca_tool/parser.py."
        )

    # Collect the header row and all data rows beneath it.
    # ASSUMPTION: data rows end at the first blank line after the header,
    # or at a row whose first cell is "End" (SimaPro's section terminator).
    data_lines: List[str] = []
    for line in lines[header_row_index:]:
        stripped = line.strip()

        if not stripped:
            # Stop at the first blank line — but only after we have at least
            # one data row (to skip blanks between the header and first row).
            if len(data_lines) > 1:
                break
            continue

        first_cell = stripped.split(delimiter)[0].strip().strip('"').lower()

        # ASSUMPTION: "End" (case-insensitive) marks the section terminator.
        if first_cell == "end":
            break

        data_lines.append(line)

    if len(data_lines) < 2:
        raise ValueError(
            "The data table header was found but contains no data rows.\n"
            "Verify that the SimaPro export includes LCIA characterisation results."
        )

    # Parse the extracted block with Python's csv module for correct handling
    # of quoted fields (SimaPro sometimes wraps names containing commas in quotes).
    reader = csv.reader(io.StringIO("\n".join(data_lines)), delimiter=delimiter)
    all_rows = [row for row in reader]

    raw_header = [cell.strip() for cell in all_rows[0]]

    # ASSUMPTION: Column layout is:
    #   [0] Impact category | [1] Unit | [2..N] One column per product system
    if len(raw_header) < 3:
        raise ValueError(
            f"Expected at least 3 columns (Impact category, Unit, one scenario), "
            f"but only found {len(raw_header)}: {raw_header}\n"
            "Check that the SimaPro export includes at least one product system."
        )

    scenario_columns = raw_header[2:]  # preserve exact scenario names from the file

    data_rows: List[Dict] = []

    for row in all_rows[1:]:
        # Pad short rows to avoid index errors (can happen on the last row).
        if len(row) < len(raw_header):
            row += [""] * (len(raw_header) - len(row))

        if not any(cell.strip() for cell in row):
            continue  # skip entirely blank rows

        impact_category = row[0].strip()
        unit = row[1].strip()

        if not impact_category:
            continue  # skip rows with no category name (e.g. subtotal rows)

        record: Optional[Dict] = {"impact_category": impact_category, "unit": unit}

        for col_idx, scenario in enumerate(scenario_columns, start=2):
            raw_value = row[col_idx].strip() if col_idx < len(row) else ""

            # Replace comma-as-decimal (European locale) with period before parsing.
            # Only applies when the delimiter is semicolon.
            if delimiter == ";" and "," in raw_value:
                raw_value = raw_value.replace(",", ".")

            try:
                # Python's float() handles scientific notation natively (e.g. "1.23E+04").
                record[scenario] = float(raw_value) if raw_value else 0.0
            except ValueError:
                # Non-numeric cell means this is a sub-header or category group row.
                # Skip it rather than failing the entire parse.
                record = None
                break

        if record is not None:
            data_rows.append(record)

    if not data_rows:
        raise ValueError(
            "No valid numeric rows were extracted from the CSV.\n"
            "The file may have an unexpected structure. "
            "Verify the export type (should be LCIA results / characterisation)."
        )

    df = pd.DataFrame(data_rows)

    # Ensure scenario columns are float dtype (not object) even if all values are 0.
    for col in scenario_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return df
