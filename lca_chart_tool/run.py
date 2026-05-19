"""
run.py — xlwings entry point for the LCA chart tool.

Two functions are called directly by Excel macro buttons:
    load_csv()       — "Load CSV" button: parse the SimaPro file, write
                       results to the Data sheet, populate scenario dropdown.
    generate_chart() — "Generate Chart" button: read Data sheet, build the
                       grouped bar chart, embed it in the Charts sheet.

Both functions wrap all logic in try/except and route all user-facing
feedback through excel_writer.write_status(). Raw Python tracebacks are
never shown to the user.

xlwings usage note:
    Both functions call xw.Book.caller() to get the workbook that triggered
    the macro. This requires the xlwings VBA module to be present in the
    workbook (see README.md for setup instructions).
"""

from pathlib import Path

import pandas as pd
import xlwings as xw

from lca_tool import excel_writer, parser, chart


def load_csv() -> None:
    """
    Parse the SimaPro CSV specified in the CSV_FILE_PATH named range and
    write the cleaned data to the Data sheet.

    Also writes discovered scenario names to SCENARIO_LIST so the
    REFERENCE_SCENARIO dropdown shows the actual scenarios from the file.

    Called by: "Load CSV" button on the Settings sheet.
    """
    wb = xw.Book.caller()
    excel_writer.write_status(wb, "Loading CSV — please wait…")

    try:
        settings = excel_writer.read_settings(wb)

        csv_path = Path(settings["csv_file_path"])

        df = parser.parse_simapro_csv(csv_path)

        scenario_columns = [
            col for col in df.columns if col not in ("impact_category", "unit")
        ]

        excel_writer.write_dataframe_to_data_sheet(wb, df)

        dropdown_warning = ""
        try:
            excel_writer.write_scenario_list(wb, scenario_columns)
        except ValueError:
            dropdown_warning = " | Note: SCENARIO_LIST range not found — Reference Scenario dropdown not updated."

        n_categories = len(df)
        n_scenarios = len(scenario_columns)
        excel_writer.write_status(
            wb,
            f"CSV loaded: {n_categories} categories, {n_scenarios} scenario(s) — "
            f"{', '.join(scenario_columns)}{dropdown_warning}"
        )

    except FileNotFoundError as exc:
        excel_writer.write_status(wb, f"File not found: {exc}")
    except ValueError as exc:
        excel_writer.write_status(wb, f"Parse error: {exc}")
    except KeyError as exc:
        excel_writer.write_status(wb, f"Settings error: {exc}")
    except Exception as exc:
        excel_writer.write_status(
            wb, f"Unexpected error during CSV load ({type(exc).__name__}): {exc}"
        )


def generate_chart() -> None:
    """
    Read parsed data from the Data sheet, apply chart settings from the
    Settings sheet, build the grouped bar chart, and embed it in the
    Charts sheet.

    Requires that load_csv() has been run first so data is present on
    the Data sheet.

    Called by: "Generate Chart" button on the Settings sheet.
    """
    wb = xw.Book.caller()
    excel_writer.write_status(wb, "Generating chart — please wait…")

    try:
        settings = excel_writer.read_settings(wb)

        # Read the data that load_csv() wrote to the Data sheet.
        data_sheet = wb.sheets[excel_writer.SHEET_DATA]
        raw = data_sheet.range("A1").expand().value

        if not raw or len(raw) < 2:
            raise ValueError(
                "No data found on the Data sheet. "
                "Click 'Load CSV' first to parse the SimaPro file."
            )

        # First row is the header, remaining rows are data.
        headers = raw[0]
        rows = raw[1:]

        df = pd.DataFrame(rows, columns=headers)

        # xlwings returns all numeric cells as floats already, but any cells
        # that were blank in Excel come back as None — coerce to 0.0.
        scenario_columns = [
            col for col in df.columns if col not in ("impact_category", "unit")
        ]
        for col in scenario_columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        fig = chart.build_chart(
            df=df,
            chart_title=settings["chart_title"] or "LCIA Results",
            value_mode=settings["value_mode"],
            bar_colors=settings["bar_colors"],
            show_data_labels=settings["show_data_labels"],
            x_label_rotation=settings["x_label_rotation"],
        )

        excel_writer.embed_chart_in_excel(wb, fig, dpi=settings["output_dpi"])

        excel_writer.write_status(
            wb, "Chart generated successfully — see the Charts sheet."
        )

    except ValueError as exc:
        excel_writer.write_status(wb, f"Chart error: {exc}")
    except KeyError as exc:
        excel_writer.write_status(wb, f"Settings error: {exc}")
    except Exception as exc:
        excel_writer.write_status(
            wb, f"Unexpected error during chart generation ({type(exc).__name__}): {exc}"
        )


if __name__ == "__main__":
    # This block only runs when the script is executed directly from a terminal,
    # not when called from Excel. Useful for verifying imports are working.
    print("run.py loaded. Functions load_csv() and generate_chart() are ready.")
    print("Run from Excel by clicking the macro buttons in the Settings sheet.")
