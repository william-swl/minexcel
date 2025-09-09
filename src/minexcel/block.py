from .utils import check_int_serial, read_excel_with_merged_cell, index_to_excel_column
import openpyxl as opx
import pandas as pd
import re
from typing import Dict, List, Tuple, Set, Any, Union, Optional


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

    tmpl_df = read_excel_with_merged_cell(template_file)

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


def parse_block(block: pd.DataFrame, tmpl: Dict[str, Any]) -> pd.DataFrame:
    """
    Parses a structured data block into a DataFrame using a template definition.

    This function processes a DataFrame according to a template specification that defines:
    - Table-level metadata locations
    - Row metadata ranges
    - Column metadata ranges
    - Data value locations
    The parsed result combines metadata with data values in long format.

    Args:
        block: Input DataFrame containing structured data to be parsed.
            Expected shape: (tmpl['block_nrow'], tmpl['block_ncol'])
        tmpl: Template dictionary defining the block structure. Expected keys:
            'block_nrow': Expected number of rows in block
            'block_ncol': Expected number of columns in block
            'tablemeta': Dict of {metadata_key: (row, col)} positions
            'rowmeta': Dict of {metadata_key: {'col': col_index, 'start': start_row, 'end': end_row}}
            'colmeta': Dict of {metadata_key: {'row': row_index, 'start': start_col, 'end': end_col}}
            'data_rows_list': List of row indices containing data values
            'data_cols_list': List of column indices containing data values

    Returns:
        pd.DataFrame: DataFrame containing:
            - 'row_index': Original row index in full valid contents (after skip rows and cols)
            - 'col_index': Original column index in full valid contents
            - 'value': Data value from the block
            - Columns for each tablemeta key with corresponding values
            - Columns for each rowmeta key with row-level metadata
            - Columns for each colmeta key with column-level metadata

    Raises:
        AssertionError: If input block dimensions don't match template specifications

    """
    # Validate input block matches template dimensions
    assert block.shape == (tmpl["block_nrow"], tmpl["block_ncol"])

    # Store original indexes for later reconstruction
    block_raw_row_idx = block.index
    block_raw_col_idx = block.columns

    # Convert to position-based integer indexes
    block = pd.DataFrame(block.values)
    # Create mapping between positional indexes and original indexes
    row_idx_map = dict(zip(block.index, block_raw_row_idx))
    col_idx_map = dict(zip(block.columns, block_raw_col_idx))

    # Extract table-level metadata
    block_tablemeta: Dict[str, Any] = {}
    if tmpl["tablemeta"]:
        for key, pos in tmpl["tablemeta"].items():
            value = block.loc[pos[0], pos[1]]
            block_tablemeta[key] = value

    # Extract row-level metadata
    block_rowmeta: Dict[str, Dict[Any, Any]] = {}
    if tmpl["rowmeta"]:
        for key, pos in tmpl["rowmeta"].items():
            # Slice row metadata range
            values = dict(block.iloc[pos["start"] : pos["end"] + 1, pos["col"]].items())
            block_rowmeta[key] = values

    # Extract column-level metadata
    block_colmeta: Dict[str, Dict[Any, Any]] = {}
    if tmpl["colmeta"]:
        for key, pos in tmpl["colmeta"].items():
            # Slice column metadata range
            values = dict(block.iloc[pos["row"], pos["start"] : pos["end"] + 1].items())
            block_colmeta[key] = values

    # Extract core data values
    data = block.iloc[tmpl["data_rows_list"], tmpl["data_cols_list"]]

    # Convert data to long format (styled structure)
    result = data.reset_index(names="row_index").melt(
        id_vars="row_index", var_name="col_index", value_name="value"
    )

    # Merge metadata into result
    for k, v in block_tablemeta.items():
        result[k] = v

    for k, d in block_rowmeta.items():
        result[k] = [d.get(i, None) for i in result["row_index"].to_list()]

    for k, d in block_colmeta.items():
        result[k] = [d.get(i, None) for i in result["col_index"].to_list()]

    # Restore original indexes
    result["row_index"] = [row_idx_map[i] for i in result["row_index"]]
    result["col_index"] = [col_idx_map[i] for i in result["col_index"]]

    return result


def read_block_excel(
    path: str,
    tmpl: Dict[str, Any],
    sheetname: Optional[str] = None,
    skipheader: int = 0,
    skipfooter: int = 0,
    skipleft: int = 0,
    skipright: int = 0,
    intervalrows: int = 0,
    intervalcols: int = 0,
) -> pd.DataFrame:
    """
    Reads structured Excel data blocks from a file and parses them according to a template.

    This function processes an Excel file containing multiple structured data blocks separated
    by specified intervals. Each block is parsed using the provided template definition to extract
    data.

    Args:
        path: Path to the Excel file
        tmpl: Template dictionary defining block structure
        sheetname: Name of the sheet to read. If None, the active sheet is used
        skipheader: Number of rows to skip at top of sheet
        skipfooter: Number of rows to skip at bottom of sheet
        skipleft: Number of columns to skip at left of sheet
        skipright: Number of columns to skip at right of sheet
        intervalrows: Number of separator rows between vertical blocks
        intervalcols: Number of separator columns between horizontal blocks

    Returns:
        pd.DataFrame: Concatenated DataFrame containing all parsed blocks in long format

    Raises:
        AssertionError: If sheet dimensions don't fit specified block structure

    """
    # Read Excel file with merged cell handling
    full = read_excel_with_merged_cell(path, sheetname=sheetname)

    # Extract full contents
    row_start = skipheader
    row_end = full.shape[0] - skipfooter
    col_start = skipleft
    col_end = full.shape[1] - skipright

    full = full.iloc[row_start:row_end, col_start:col_end]

    # Reset index to integer-based for consistent slicing
    full = pd.DataFrame(full.values)

    # Calculate number of horizontal blocks
    nblock_each_row = round(full.shape[1] / (tmpl["block_ncol"] + intervalcols))
    expected_cols = (
        nblock_each_row * tmpl["block_ncol"] + (nblock_each_row - 1) * intervalcols
    )
    full_content_cols = full.shape[1]

    assert full_content_cols == expected_cols, (
        f"Columns don't fit block: {expected_cols} cols expected, {full_content_cols} cols detected in valid full content. Check your skip* config"
    )

    # Calculate number of vertical blocks
    nblock_each_col = round(full.shape[0] / (tmpl["block_nrow"] + intervalrows))
    expected_rows = (
        nblock_each_col * tmpl["block_nrow"] + (nblock_each_col - 1) * intervalrows
    )
    full_content_rows = full.shape[0]

    assert full_content_rows == expected_rows, (
        f"Rows don't fit block: {expected_rows} rows expected, {full_content_rows} rows detected in valid full content. Check your skip* config"
    )

    # Calculate starting positions for each block
    block_row_starts = [
        i * tmpl["block_nrow"] + i * intervalrows for i in range(nblock_each_col)
    ]
    block_col_starts = [
        i * tmpl["block_ncol"] + i * intervalcols for i in range(nblock_each_row)
    ]

    # Extract all valid data blocks
    blocks = []
    for row_start in block_row_starts:
        for col_start in block_col_starts:
            # Extract block using template dimensions
            block = full.iloc[
                range(row_start, row_start + tmpl["block_nrow"]),
                range(col_start, col_start + tmpl["block_ncol"]),
            ]
            # Skip fully empty blocks
            if all(block.isna().all()):
                continue
            blocks.append(block)

    # Parse each block using template and combine results
    result = [parse_block(b, tmpl) for b in blocks]
    result = pd.concat(result, axis=0)

    result["row_excel"] = [i + skipheader + 1 for i in result["row_index"]]
    result["col_excel"] = [
        index_to_excel_column(i + skipleft) for i in result["col_index"]
    ]

    return result
