from .block import read_block, parse_template
from .utils import check_int_serial


def main() -> None:
    print("Hello from minexcel!")


__all__ = ["read_block", "check_int_serial", "parse_template"]
