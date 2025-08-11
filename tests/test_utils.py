import minexcel as mxl


def test_check_int_serial():
    assert mxl.check_int_serial([1, 2, 3])
    assert mxl.check_int_serial([1, 3, 3]) is False
    assert mxl.check_int_serial([1, 3, 2, 4]) is False
    assert mxl.check_int_serial([1, 3, 2, 4], sort=True)
