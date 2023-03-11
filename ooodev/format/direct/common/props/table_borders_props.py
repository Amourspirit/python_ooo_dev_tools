from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class TableBordersProps(NamedTuple):
    tbl_border: str
    shadow: str
    tbl_distance: str
    merge: str
