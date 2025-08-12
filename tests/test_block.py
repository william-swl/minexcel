import minexcel as mxl


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
