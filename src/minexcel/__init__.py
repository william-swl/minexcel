from .block import read_block, parse_template, parse_block
from .utils import check_int_serial, read_excel_with_merged_cell


def main() -> None:
    print("Hello from minexcel!")


__all__ = [
    "read_block",
    "parse_block",
    "check_int_serial",
    "parse_template",
    "read_excel_with_merged_cell",
]
