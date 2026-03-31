import matplotlib.font_manager as font_manager
import matplotlib.pyplot as plt
import pandas as pd
from pandas.plotting import table
import re
import numpy as np
from typing import Any


def save_df_as_image(df: pd.DataFrame, filename: str) -> None:
    """
    Render a pandas DataFrame as a styled table and save it as an image.

    The function uses matplotlib to draw the DataFrame as a table on a figure,
    hides axes, applies basic styling (font, background, header color),
    and saves the result to disk as a PNG (or other supported image format).

    :param df: DataFrame to render as an image
    :type df: pandas.DataFrame
    :param filename: Output file path where the image will be saved
    :type filename: str
    :return: None
    :rtype: None
    """
    plt.rcParams['font.size'] = 32
    plt.rcParams['font.family'] = 'DejaVu Sans'
    plt.rcParams['font.weight'] = 'bold'

    fig_width: float = len(df.columns) * 2
    fig_height: float = len(df) * 0.2
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    fig.patch.set_facecolor('lightgrey')

    ax.axis('tight')
    ax.axis('off')

    df_for_display: pd.DataFrame = df.copy()
    df_for_display.index = [''] * len(df)

    col_widths = [0.7, 0.15, 0.15]
    _ = table(
        ax,
        df_for_display,
        loc='center',
        cellLoc='center',
        colWidths=col_widths,
        colColours=["#FFD700"] * len(df.columns),
    )

    plt.subplots_adjust(left=0.05, right=0.95, top=0.99, bottom=0.01)

    plt.savefig(filename, dpi=300, bbox_inches='tight', pad_inches=0)
    plt.close()


def extract_year(title: str) -> float:
    """
    Extract the first 4-digit year (1900–2099) from a string.

    :param title: Input string, typically a movie title possibly containing a year
    :type title: str
    :return: Extracted year as integer if found, otherwise NaN
    :rtype: float
    """
    matches = re.findall(r'\b(19|20)\d{2}\b', title)
    if matches:
        return int(matches[-1])
    return float("nan")


def define_decade(year: float) -> float:
    """
    Map a given year to the start of its decade.

    For example, 1994 -> 1990, 2001 -> 2000. If the input is NaN,
    the function returns NaN.

    :param year: Calendar year or NaN
    :type year: float
    :return: Decade start year (e.g. 1990, 2000) or NaN
    :rtype: float
    """
    if not np.isnan(year):
        return (int(year) // 10) * 10
    return float("nan")


def process_results(subset: Any) -> pd.DataFrame:
    """
    Aggregate ballot results for a subset of votes.

    The input subset must be convertible to a DataFrame with at least the
    following columns: ``'Фильм'``, ``'Упоминаний'``, ``'Баллов'``, ``'Места'``.
    The function groups votes by film, sums scores and mentions, collects
    all positions into a list, and returns a sorted table.

    :param subset: Iterable or DataFrame-like object with voting data
    :type subset: Any
    :return: Processed results with total scores, mentions and position tuples
    :rtype: pandas.DataFrame
    """
    top_subset: pd.DataFrame = pd.DataFrame(subset)

    grouped_top_subset: pd.DataFrame = top_subset.groupby('Фильм').agg({
        'Упоминаний': 'sum',
        'Баллов': 'sum'
    }).reset_index()

    grouped_positions_subset: pd.DataFrame = (
        top_subset.groupby('Фильм')['Места'].apply(list).reset_index()
    )

    merged_group_subset: pd.DataFrame = pd.merge(
        grouped_top_subset,
        grouped_positions_subset,
        on='Фильм'
    )

    sorted_group_subset: pd.DataFrame = merged_group_subset.sort_values(
        by=['Баллов', 'Упоминаний'],
        ascending=[False, False]
    )
    sorted_group_subset['Позиции в списках'] = sorted_group_subset['Места'].apply(tuple)

    return sorted_group_subset.drop(columns=['Места'])
