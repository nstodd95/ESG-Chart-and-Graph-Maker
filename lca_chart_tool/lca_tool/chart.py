"""
chart.py — Matplotlib grouped bar chart generator for LCIA results.

KNOWN LIMITATION — single Y-axis and order-of-magnitude differences:
    LCIA impact categories routinely span 10–20 orders of magnitude across
    categories (e.g. Global Warming in kg CO2-eq vs. Ionising Radiation in
    kBq U235-eq). Plotting all on one chart with a shared Y-axis makes most
    bars invisible or misleading.

    This module resolves the problem in two ways:
        1. "Normalized (% of max)" mode (RECOMMENDED): each row is scaled to
           the percentage of its own maximum value across scenarios. This
           removes the cross-category scale problem entirely and allows
           meaningful visual comparison of scenario trade-offs.
        2. "Raw values" mode: the chart is drawn with raw units, but a
           prominent warning annotation is added so users understand that
           bars are NOT visually comparable across categories. Raw mode is
           appropriate only for single-category or same-unit exports.
"""

from typing import Dict, List, Optional, Tuple

import matplotlib
# Use Agg (non-interactive) backend so matplotlib never tries to open a GUI
# window. This is required for operation inside xlwings on both platforms.
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.ticker
import numpy as np
import pandas as pd


_DEFAULT_FIGURE_SIZE: Tuple[float, float] = (16, 7)
_DEFAULT_COLOR = "#4472C4"

_RAW_VALUES_WARNING = (
    "WARNING: Raw values plotted on a single Y-axis.\n"
    "Impact categories span many orders of magnitude — bars are NOT\n"
    "comparable across categories. Use 'Normalized (% of max)' for\n"
    "visual comparison of scenario trade-offs."
)


def _get_scenario_columns(df: pd.DataFrame) -> List[str]:
    """Return the list of scenario column names (all columns except metadata)."""
    return [col for col in df.columns if col not in ("impact_category", "unit")]


def _normalize_to_percent_of_max(
    df: pd.DataFrame, scenario_columns: List[str]
) -> pd.DataFrame:
    """
    Normalise scenario values within each impact category row to percentage
    of the maximum absolute value in that row.

    A category where all scenario values are zero is left as all-zeros to
    avoid division by zero.

    Args:
        df: DataFrame with columns [impact_category, unit, scenario_1, ...].
        scenario_columns: Names of the scenario columns to normalise.

    Returns:
        A copy of df with scenario values replaced by percentages (0–100).
    """
    df_norm = df.copy()
    for idx in df_norm.index:
        row_vals = df_norm.loc[idx, scenario_columns]
        row_max = row_vals.abs().max()
        if row_max != 0:
            df_norm.loc[idx, scenario_columns] = row_vals / row_max * 100
    return df_norm


def _resolve_colors(bar_colors: List[str], n_scenarios: int) -> List[str]:
    """
    Build the list of bar colors, cycling through bar_colors if there are
    more scenarios than colors provided. Falls back to _DEFAULT_COLOR for
    any slot that contains an invalid or empty string.

    Args:
        bar_colors: List of hex color strings from settings (e.g. ["#4472C4", ...]).
        n_scenarios: Number of scenario columns to color.

    Returns:
        A list of exactly n_scenarios color strings.
    """
    resolved: List[str] = []
    for i in range(n_scenarios):
        raw = bar_colors[i % len(bar_colors)] if bar_colors else _DEFAULT_COLOR
        color = raw.strip() if raw else _DEFAULT_COLOR
        if not color:
            color = _DEFAULT_COLOR
        if not color.startswith("#"):
            color = "#" + color
        resolved.append(color)
    return resolved


def _add_data_labels(ax: plt.Axes, bars, normalized: bool) -> None:
    """
    Annotate each bar with its numeric value.

    Bars with a height of zero are skipped to avoid clutter.

    Args:
        ax: The Axes object containing the bars.
        bars: BarContainer returned by ax.bar().
        normalized: If True, format values as "XX.X%" else as scientific notation.
    """
    for bar in bars:
        height = bar.get_height()
        if height == 0:
            continue
        label = f"{height:.1f}%" if normalized else f"{height:.3g}"
        ax.annotate(
            label,
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 3),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=6.5,
            clip_on=True,
        )


def _build_stacked_percent_chart(
    df: pd.DataFrame,
    chart_title: str,
    bar_colors: List[str],
    show_data_labels: bool,
    x_label_rotation: int,
    figure_size: Tuple[float, float],
) -> plt.Figure:
    """
    Build a 100% stacked bar chart where each component is shown as a
    percentage of the "Total" column for that impact category.

    One bar per impact category. Each bar is divided into segments, one per
    component scenario. The "Total" column is used as the denominator and is
    not drawn as its own segment.

    If no column named "Total" (case-insensitive) exists, the denominator
    falls back to the sum of all scenario values in that row.

    Args:
        df: Cleaned DataFrame from parser.parse_simapro_csv().
        chart_title: Title string displayed at the top of the chart.
        bar_colors: List of hex color strings, one per component scenario.
        show_data_labels: If True, label each segment with its percentage
            (only drawn when the segment is wide enough to be readable).
        x_label_rotation: Degrees to rotate x-axis category labels.
        figure_size: (width_inches, height_inches) for the figure.

    Returns:
        A matplotlib Figure object ready for embedding into Excel.
    """
    scenario_columns = _get_scenario_columns(df)

    # Find the "Total" column to use as denominator.
    total_col = next(
        (col for col in scenario_columns if col.strip().lower() == "total"),
        None,
    )

    # Segments are all scenario columns except "Total" itself.
    stack_cols = [col for col in scenario_columns if col != total_col]

    if not stack_cols:
        raise ValueError(
            "No component columns found to stack. "
            "After excluding the 'Total' column, at least one scenario column is required."
        )

    # Calculate each component's share of the total for each row.
    plot_df = df.copy()
    if total_col is not None:
        denominators = plot_df[total_col].abs()
    else:
        # No Total column: use row sum as denominator.
        denominators = plot_df[stack_cols].abs().sum(axis=1)

    for col in stack_cols:
        plot_df[col] = plot_df.apply(
            lambda row, c=col: (
                abs(row[c]) / denominators[row.name] * 100
                if denominators[row.name] != 0 else 0.0
            ),
            axis=1,
        )

    x_labels: List[str] = [
        f"{row['impact_category']}\n[{row['unit']}]"
        for _, row in plot_df.iterrows()
    ]

    n_categories = len(plot_df)
    x_positions = np.arange(n_categories, dtype=float)
    effective_colors = _resolve_colors(bar_colors, len(stack_cols))

    fig, ax = plt.subplots(figsize=figure_size)

    bottoms = np.zeros(n_categories)

    for col, color in zip(stack_cols, effective_colors):
        values = plot_df[col].to_numpy(dtype=float)
        bars = ax.bar(
            x_positions,
            values,
            bottom=bottoms,
            label=col,
            color=color,
            edgecolor="white",
            linewidth=0.5,
        )

        if show_data_labels:
            for bar, val in zip(bars, values):
                # Only label segments large enough to be readable.
                if val >= 4:
                    ax.text(
                        bar.get_x() + bar.get_width() / 2,
                        bar.get_y() + bar.get_height() / 2,
                        f"{val:.1f}%",
                        ha="center",
                        va="center",
                        fontsize=6.5,
                        color="white",
                        fontweight="bold",
                    )

        bottoms += values

    ax.set_xticks(x_positions)
    ax.set_xticklabels(
        x_labels,
        rotation=x_label_rotation,
        ha="right" if x_label_rotation > 0 else "center",
        fontsize=8,
    )
    ax.set_ylabel("% of Total", fontsize=10)
    ax.set_ylim(0, 110)
    ax.yaxis.set_major_formatter(
        matplotlib.ticker.FuncFormatter(lambda val, _: f"{val:.0f}%")
    )
    ax.axhline(100, color="black", linewidth=0.8, linestyle="--", alpha=0.4)
    ax.set_title(chart_title, fontsize=14, fontweight="bold", pad=14)

    # Legend placed outside the plot area to avoid overlapping tall stacks.
    ax.legend(
        loc="upper left",
        bbox_to_anchor=(1.01, 1),
        framealpha=0.9,
        fontsize=8,
        borderaxespad=0,
    )

    fig.tight_layout()
    return fig


def build_chart(
    df: pd.DataFrame,
    chart_title: str,
    value_mode: str,
    bar_colors: List[str],
    show_data_labels: bool,
    x_label_rotation: int,
    figure_size: Tuple[float, float] = _DEFAULT_FIGURE_SIZE,
) -> plt.Figure:
    """
    Build a chart from a parsed SimaPro LCIA DataFrame.

    Dispatches to the appropriate chart type based on value_mode:
        "Stacked (% of Total)"  — 100% stacked bar, components as % of Total
        "Normalized (% of max)" — grouped bar, each category scaled to its own max
        "Raw values"            — grouped bar, raw units with cross-category warning

    Args:
        df: Cleaned DataFrame from parser.parse_simapro_csv(). Must have
            columns: impact_category, unit, and at least one scenario column.
        chart_title: Title string displayed at the top of the chart.
        value_mode: One of the three mode strings above.
        bar_colors: List of hex strings, one per scenario. Cycles if there
            are more scenarios than colors.
        show_data_labels: If True, annotate bars with their numeric values.
        x_label_rotation: Degrees to rotate x-axis category labels.
        figure_size: (width_inches, height_inches) for the figure.

    Returns:
        A matplotlib Figure object ready for embedding into Excel.

    Raises:
        ValueError: If the DataFrame is empty or has no scenario columns.
    """
    if df.empty:
        raise ValueError("Cannot build a chart from an empty DataFrame.")

    scenario_columns = _get_scenario_columns(df)

    if not scenario_columns:
        raise ValueError(
            "No scenario columns found. The parsed data must contain at least "
            "one product system / scenario column beyond 'Impact category' and 'Unit'."
        )

    mode = value_mode.strip().lower()

    # Dispatch to stacked chart.
    if "stacked" in mode:
        return _build_stacked_percent_chart(
            df, chart_title, bar_colors, show_data_labels, x_label_rotation, figure_size
        )

    # Grouped bar chart — normalized or raw.
    normalized = "normalized" in mode

    if normalized:
        plot_df = _normalize_to_percent_of_max(df, scenario_columns)
        y_label = "% of maximum value within each category"
    else:
        plot_df = df.copy()
        y_label = "Raw value (shared Y-axis — categories NOT comparable)"

    x_labels: List[str] = [
        f"{row['impact_category']}\n[{row['unit']}]"
        for _, row in plot_df.iterrows()
    ]

    n_categories = len(plot_df)
    n_scenarios = len(scenario_columns)
    group_width = 0.8
    bar_width = group_width / n_scenarios
    x_positions = np.arange(n_categories, dtype=float)
    effective_colors = _resolve_colors(bar_colors, n_scenarios)

    fig, ax = plt.subplots(figsize=figure_size)

    for i, (scenario, color) in enumerate(zip(scenario_columns, effective_colors)):
        offsets = x_positions + (i - n_scenarios / 2 + 0.5) * bar_width
        values = plot_df[scenario].to_numpy(dtype=float)
        bars = ax.bar(
            offsets, values, width=bar_width, label=scenario,
            color=color, edgecolor="white", linewidth=0.5,
        )
        if show_data_labels:
            _add_data_labels(ax, bars, normalized)

    ax.set_xticks(x_positions)
    ax.set_xticklabels(
        x_labels,
        rotation=x_label_rotation,
        ha="right" if x_label_rotation > 0 else "center",
        fontsize=8,
    )
    ax.set_ylabel(y_label, fontsize=10)
    ax.set_title(chart_title, fontsize=14, fontweight="bold", pad=14)
    ax.legend(loc="upper right", framealpha=0.9, fontsize=9)
    ax.axhline(0, color="black", linewidth=0.6, zorder=0)

    if normalized:
        ax.set_ylim(0, 110)
        ax.yaxis.set_major_formatter(
            matplotlib.ticker.FuncFormatter(lambda val, _: f"{val:.0f}%")
        )

    if not normalized:
        ax.annotate(
            _RAW_VALUES_WARNING,
            xy=(0.01, 0.99), xycoords="axes fraction",
            fontsize=7, color="#8B0000", va="top", ha="left",
            bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow",
                      edgecolor="orange", alpha=0.9),
        )

    fig.tight_layout()
    return fig
