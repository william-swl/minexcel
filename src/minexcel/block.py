from .utils import check_int_serial
import openpyxl as opx
import pandas as pd
import re
from typing import Dict, List, Tuple, Set, Any, Union


def parse_template(template_file: str) -> Dict[str, Any]:
    """
    Parses an Excel template file to extract metadata information about its structure.

    This function analyzes an Excel template and returns a dictionary containing:
    - Dimensions of the entire template and data blocks
    - Positions of row and column metadata markers
    - Table-level metadata positions
    - Validation checks for template structure

    Args:
        template_file: Path to the Excel template file (.xlsx) to be parsed

    Returns:
        A dictionary with the following structure:

        - 'block_nrow': int,          # Total rows in template
        - 'block_ncol': int,          # Total columns in template
        - 'data_nrow': int,           # Number of rows in data zone
        - 'data_ncol': int,           # Number of columns in data zone
        - 'data_rows_list': List[int], # Row indices in data zone
        - 'data_cols_list': List[int], # Column indices in data zone
        - 'tablemeta': Dict[str, Tuple[int, int]], # Table metadata positions
        - 'rowmeta': Dict[str, Dict[str, int]],    # Row metadata definitions
        - 'colmeta': Dict[str, Dict[str, int]]     # Column metadata definitions

    Raises:
        AssertionError: If the template structure doesn't meet requirements:

            - Data zone must be contiguous
            - Each rowmeta key must appear only in a single column
            - Each colmeta key must appear only in a single row
            - Metadata ranges must be serial and within data zone

    """
    result = {
        "block_nrow": None,
        "block_ncol": None,
        "data_nrow": None,
        "data_ncol": None,
        "data_rows_list": None,
        "data_cols_list": None,
        "tablemeta": {},
        "rowmeta": {},
        "colmeta": {},
    }

    # Load workbook and get active sheet
    workbook = opx.load_workbook(template_file)
    template = workbook.active

    # Process merged cells by unmerging and copying top-left value to all cells
    merged_ranges = [m for m in template.merged_cells.ranges]

    for merged_range in merged_ranges:
        template.unmerge_cells(str(merged_range))

    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        cell_value = template.cell(row=min_row, column=min_col).value

        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                template.cell(row=row, column=col).value = cell_value

    # Convert to DataFrame for easier processing
    tmpl_df = pd.DataFrame(template.values, columns=None)

    # 1. Get basic dimensions
    result["block_nrow"], result["block_ncol"] = tmpl_df.shape

    # 2. Identify data zone (non-blank areas)
    data_cols = tmpl_df.isna().sum(axis=0) != 0
    data_cols_list = data_cols[data_cols].index.to_list()

    data_rows = tmpl_df.isna().sum(axis=1) != 0
    data_rows_list = data_rows[data_rows].index.to_list()

    # 3. Validate data zone is contiguous
    assert check_int_serial(data_cols_list), "Data columns must be contiguous"
    assert check_int_serial(data_rows_list), "Data rows must be contiguous"

    result["data_cols_list"] = data_cols_list
    result["data_rows_list"] = data_rows_list
    result["data_nrow"], result["data_ncol"] = tmpl_df.loc[data_rows, data_cols].shape

    # 4. Extract metadata
    # Table-level metadata (can appear anywhere)
    for irow, row in tmpl_df.iterrows():
        for icol, cell in enumerate(row):
            if not isinstance(cell, str):
                continue
            match_res = re.match(r"(\w+)\[tablemeta\]", str(cell))
            if match_res:
                key = match_res.group(1)
                result["tablemeta"].setdefault(key, []).append((irow, icol))

    # Each tablemeta key should have only one position
    for key in list(result["tablemeta"].keys()):
        positions = result["tablemeta"][key]
        if len(positions) > 1:
            # Use the first found position
            result["tablemeta"][key] = positions[0]

    # Row metadata (must appear in single columns)
    rowmeta_df = tmpl_df.loc[:, ~data_cols]  # Metadata columns (non-data)

    for icol, col in rowmeta_df.items():  # Iterate by column
        key_list = []
        range_list = []
        for irow, cell in col.items():
            if not isinstance(cell, str):
                continue
            match_res = re.match(r"(\w+)\[rowmeta\]", str(cell))
            if match_res:
                key_list.append(match_res.group(1))
                range_list.append(irow)

        if key_list:
            # Ensure single key per column and contiguous range
            unique_keys = set(key_list)
            assert len(unique_keys) == 1, (
                f"Multiple rowmeta keys found in column {icol}: {unique_keys}"
            )

            meta_key = next(iter(unique_keys))
            assert check_int_serial(range_list), (
                f"Rowmeta ranges must be contiguous in column {icol}"
            )
            assert set(range_list).issubset(set(data_rows_list)), (
                f"Rowmeta rows out of data zone in column {icol}"
            )

            result["rowmeta"][meta_key] = {
                "col": icol,
                "start": min(range_list),
                "end": max(range_list),
            }

    # Column metadata (must appear in single rows)
    colmeta_df = tmpl_df.loc[~data_rows, :]  # Metadata rows (non-data)

    for irow, row in colmeta_df.iterrows():  # Iterate by row
        key_list = []
        range_list = []
        for icol, cell in enumerate(row):
            if not isinstance(cell, str):
                continue
            match_res = re.match(r"(\w+)\[colmeta\]", str(cell))
            if match_res:
                key_list.append(match_res.group(1))
                range_list.append(icol)

        if key_list:
            # Ensure single key per row and contiguous range
            unique_keys = set(key_list)
            assert len(unique_keys) == 1, (
                f"Multiple colmeta keys found in row {irow}: {unique_keys}"
            )

            meta_key = next(iter(unique_keys))
            assert check_int_serial(range_list), (
                f"Colmeta ranges must be contiguous in row {irow}"
            )
            assert set(range_list).issubset(set(data_cols_list)), (
                f"Colmeta columns out of data zone in row {irow}"
            )

            result["colmeta"][meta_key] = {
                "row": irow,
                "start": min(range_list),
                "end": max(range_list),
            }

    return result


def read_block():
    return 0
