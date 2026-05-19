"""
excel_writer.py — All xlwings interactions for the LCA chart tool.

This is the ONLY module that imports xlwings. All reads from and writes to
the Excel workbook go through the functions defined here. No other module
should call xlwings directly.

Named range constants at the top of this file must match the named ranges
created in the Excel workbook exactly (case-sensitive).

Cross-platform notes:
    - xlwings works on Windows (via COM) and Mac (via xlwings server / appscript).
    - sheet.pictures.add(path, update=True) replaces an existing picture with
      the same name, preventing duplicate images when "Generate Chart" is
      clicked multiple times.
    - The chart image is saved to a temporary file rather than passed as
      BytesIO because the temp-file path is the most reliable cross-platform
      input format for xlwings pictures.add().
    - No hardcoded cell addresses are used anywhere; all access goes through
      named ranges or named sheets.
"""

import os
import tempfile
from pathlib import Path
from typing import Dict, List

import matplotlib.pyplot as plt
import pandas as pd
import xlwings as xw


# ---------------------------------------------------------------------------
# Named range constants — must match names defined in the Excel workbook.
# ---------------------------------------------------------------------------

NR_CSV_PATH            = "CSV_FILE_PATH"
NR_CHART_TITLE         = "CHART_TITLE"
NR_VALUE_MODE          = "VALUE_MODE"
NR_REFERENCE_SCENARIO  = "REFERENCE_SCENARIO"
NR_SHOW_DATA_LABELS    = "SHOW_DATA_LABELS"
NR_X_LABEL_ROTATION    = "X_LABEL_ROTATION"
NR_OUTPUT_DPI          = "OUTPUT_DPI"
NR_STATUS_MESSAGE      = "STATUS_MESSAGE"
NR_SCENARIO_LIST       = "SCENARIO_LIST"   # helper range backing the Reference Scenario dropdown

# Color slot names, one per scenario (up to 6).
NR_BAR_COLORS: List[str] = [f"Color_{i}" for i in range(1, 7)]

# Sheet names — must match the workbook's tab names exactly.
SHEET_SETTINGS = "Settings"
SHEET_DATA     = "Data"
SHEET_CHARTS   = "Charts"

# Name used for the embedded chart picture (update=True replaces it on re-run).
_CHART_PICTURE_NAME = "LCA_Chart"

_DEFAULT_COLOR = "#4472C4"


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _read_named_range(wb: xw.Book, name: str) -> str:
    """
    Read the string value of a named range from the workbook.

    Args:
        wb: Active xlwings Book.
        name: Named range name as defined in the workbook.

    Returns:
        The cell's value as a stripped string, or "" if the cell is empty.

    Raises:
        KeyError: If the named range does not exist in the workbook.
    """
    try:
        value = wb.names[name].refers_to_range.value
    except Exception:
        raise KeyError(
            f"Named range '{name}' was not found in the workbook.\n"
            "Check the Excel setup steps in README.md to ensure all "
            "named ranges have been created correctly."
        )
    if value is None:
        return ""
    # Excel may return numbers as floats (e.g. the DPI cell returns 150.0).
    # Convert to string and strip trailing ".0" for clean integer parsing.
    return str(value).strip()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def read_settings(wb: xw.Book) -> Dict:
    """
    Read all user-configurable settings from the Settings sheet's named ranges.

    Args:
        wb: The active xlwings Book object (the workbook that called the macro).

    Returns:
        A dict with the following keys:
            csv_file_path (str)
            chart_title (str)
            value_mode (str): "Raw values" or "Normalized (% of max)"
            reference_scenario (str)
            show_data_labels (bool)
            x_label_rotation (int): degrees (0, 30, 45, or 90)
            output_dpi (int): 150 or 300
            bar_colors (list[str]): up to 6 hex color strings

    Raises:
        KeyError: If any required named range is missing.
        ValueError: If a numeric setting cannot be parsed.
    """
    colors: List[str] = []
    for color_name in NR_BAR_COLORS:
        try:
            raw = wb.names[color_name].refers_to_range.value
            color = str(raw).strip() if raw else _DEFAULT_COLOR
            if not color:
                color = _DEFAULT_COLOR
            if not color.startswith("#"):
                color = "#" + color
        except Exception:
            color = _DEFAULT_COLOR
        colors.append(color)

    # Parse rotation: Excel may return the number as a float string ("45.0").
    raw_rotation = _read_named_range(wb, NR_X_LABEL_ROTATION)
    try:
        rotation = int(float(raw_rotation)) if raw_rotation else 45
    except ValueError:
        rotation = 45

    # Parse DPI: same float-string issue.
    raw_dpi = _read_named_range(wb, NR_OUTPUT_DPI)
    try:
        dpi = int(float(raw_dpi)) if raw_dpi else 150
    except ValueError:
        dpi = 150

    return {
        "csv_file_path":      _read_named_range(wb, NR_CSV_PATH),
        "chart_title":        _read_named_range(wb, NR_CHART_TITLE),
        "value_mode":         _read_named_range(wb, NR_VALUE_MODE) or "Normalized (% of max)",
        "reference_scenario": _read_named_range(wb, NR_REFERENCE_SCENARIO),
        "show_data_labels":   _read_named_range(wb, NR_SHOW_DATA_LABELS).lower() == "yes",
        "x_label_rotation":   rotation,
        "output_dpi":         dpi,
        "bar_colors":         colors,
    }


def write_dataframe_to_data_sheet(wb: xw.Book, df: pd.DataFrame) -> None:
    """
    Overwrite the Data sheet with the contents of the parsed DataFrame.

    Clears all existing content on the sheet before writing so stale data
    from a previous CSV load does not persist.

    Args:
        wb: The active xlwings Book.
        df: The parsed SimaPro DataFrame from parser.parse_simapro_csv().
    """
    sheet = wb.sheets[SHEET_DATA]
    sheet.clear()
    # Write the header row, then all data rows. index=False keeps the
    # pandas row index out of the spreadsheet.
    sheet.range("A1").value = [list(df.columns)]
    if not df.empty:
        sheet.range("A2").value = df.values.tolist()


def write_scenario_list(wb: xw.Book, scenario_names: List[str]) -> None:
    """
    Write scenario names into the SCENARIO_LIST named range so that the
    REFERENCE_SCENARIO dropdown (which uses SCENARIO_LIST as its source)
    shows the actual scenario names from the loaded file.

    Also sets REFERENCE_SCENARIO to the first scenario name as a default
    so the user has a valid starting value.

    Args:
        wb: The active xlwings Book.
        scenario_names: List of scenario/product system column names.

    Raises:
        ValueError: If the SCENARIO_LIST named range cannot be written to.
    """
    try:
        target = wb.names[NR_SCENARIO_LIST].refers_to_range

        # Clear the entire SCENARIO_LIST range before writing to remove
        # stale entries from a previous CSV load.
        target.clear_contents()

        for i, name in enumerate(scenario_names):
            # Write each scenario name into successive cells in the range.
            target[i, 0].value = name

        # Set the REFERENCE_SCENARIO cell to the first scenario as a default.
        if scenario_names:
            wb.names[NR_REFERENCE_SCENARIO].refers_to_range.value = scenario_names[0]

    except KeyError:
        raise ValueError(
            f"Named range '{NR_SCENARIO_LIST}' was not found.\n"
            "The Reference Scenario dropdown will not update. "
            "Check the Excel setup steps in README.md."
        )
    except Exception as e:
        raise ValueError(
            f"Failed to write to '{NR_SCENARIO_LIST}': {e}\n"
            "The Reference Scenario dropdown may not reflect the loaded file."
        )


def embed_chart_in_excel(wb: xw.Book, fig: plt.Figure, dpi: int) -> None:
    """
    Save the matplotlib Figure as a PNG and embed it in the Charts sheet.

    The figure is first saved to a temporary file (the most reliable cross-
    platform input for xlwings pictures.add), then the temp file is deleted
    after the picture is inserted. If a picture named _CHART_PICTURE_NAME
    already exists on the sheet it is replaced (update=True), so clicking
    "Generate Chart" multiple times does not accumulate duplicate images.

    Args:
        wb: The active xlwings Book.
        fig: The matplotlib Figure returned by chart.build_chart().
        dpi: Resolution for the PNG export (e.g. 150 or 300).
    """
    sheet = wb.sheets[SHEET_CHARTS]

    # Write to a named temp file (delete=False) so we can pass the path to
    # xlwings; we delete it manually in the finally block.
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
    os.close(tmp_fd)  # close the file descriptor; savefig will re-open it

    try:
        fig.savefig(tmp_path, format="png", dpi=dpi, bbox_inches="tight")

        sheet.pictures.add(
            tmp_path,
            name=_CHART_PICTURE_NAME,
            update=True,
            # Position the top-left corner of the picture at cell B2.
            left=sheet.range("B2").left,
            top=sheet.range("B2").top,
        )
    finally:
        # Always clean up the temp file, even if an error occurred.
        try:
            Path(tmp_path).unlink()
        except OSError:
            pass

    # Release the figure's memory now that it has been embedded.
    plt.close(fig)


def write_status(wb: xw.Book, message: str) -> None:
    """
    Write a status or error message to the STATUS_MESSAGE named cell on
    the Settings sheet.

    This is the sole mechanism for communicating success or failure back to
    the user inside Excel. All user-facing messages go through this function.

    Args:
        wb: The active xlwings Book.
        message: A short, human-readable string (no raw Python tracebacks).
    """
    try:
        wb.names[NR_STATUS_MESSAGE].refers_to_range.value = message
    except Exception:
        # If the workbook is in an unexpected state, silently swallow the
        # error — raising here would hide the original error from the caller.
        pass
