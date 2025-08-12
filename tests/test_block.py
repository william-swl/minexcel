import minexcel as mxl
import pandas as pd

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)


def test_parse_template():
    path = "tests/files/template.xlsx"
    res = mxl.parse_template(path)
    res_ref = {
        "block_nrow": 8,
        "block_ncol": 12,
        "data_nrow": 6,
        "data_ncol": 10,
        "data_rows_list": [2, 3, 4, 5, 6, 7],
        "data_cols_list": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        "tablemeta": {"batch": (0, 0)},
        "rowmeta": {
            "item": {"col": 0, "start": 2, "end": 7},
            "titer": {"col": 11, "start": 2, "end": 7},
        },
        "colmeta": {
            "conc": {"row": 0, "start": 2, "end": 10},
            "ab": {"row": 1, "start": 1, "end": 10},
        },
    }

    assert res == res_ref


def test_read_block_excel0(snapshot):
    tmpl = mxl.parse_template("tests/files/template.xlsx")
    path = "tests/files/data.xlsx"
    res = mxl.read_block_excel(path, tmpl, skipheader=1)
    snapshot.assert_match(res)


def test_read_block_excel1(snapshot):
    tmpl = mxl.parse_template("tests/files/template.xlsx")
    path = "tests/files/data1.xlsx"
    res = mxl.read_block_excel(path, tmpl, skipheader=1, intervalcols=1)
    snapshot.assert_match(res)


def test_read_block_excel2(snapshot):
    tmpl = mxl.parse_template("tests/files/template.xlsx")
    path = "tests/files/data2.xlsx"
    res = mxl.read_block_excel(path, tmpl, skipheader=1, intervalrows=2, intervalcols=1)
    snapshot.assert_match(res)
