from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class TableBordersProps(NamedTuple):
    """Internal Properties"""

    tbl_border: str
    shadow: str
    tbl_distance: str
    merge: str
    tbl_bdr_left: PropPair
    tbl_bdr_top: PropPair
    tbl_bdr_right: PropPair
    tbl_bdr_bottom: PropPair
    tbl_bdr_horz: PropPair
    tbl_bdr_vert: PropPair
    tbl_bdr_dist: PropPair
