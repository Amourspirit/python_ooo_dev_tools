from __future__ import annotations
from typing import NamedTuple


class CellTextAlignProps(NamedTuple):
    """Internal Properties"""

    hori_justify: str
    hori_method: str
    vert_justify: str
    vert_method: str
    indent: str
