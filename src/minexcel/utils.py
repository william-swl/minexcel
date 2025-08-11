def check_int_serial(ser, sort=False):
    # sort

    if sort:
        ser = sorted(list(ser))
    else:
        ser = list(ser)

    range_list = list(range(min(ser), max(ser) + 1))

    return ser == range_list
