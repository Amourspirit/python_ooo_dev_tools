from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class CellBordersProps(NamedTuple):
    """Internal Properties"""

    tbl_border: str  # TableBorder2
    shadow: str  # ShadowFormat
    diag_up: str  # DiagonalBLTR2
    diag_dn: str  # DiagonalTLBR2
    pad_left: str  # ParaLeftMargin
    pad_top: str  # ParaTopMargin
    pad_right: str  # ParaRightMargin
    pad_btm: str  # ParaBottomMargin
    tbl_bdr_left: PropPair
    tbl_bdr_top: PropPair
    tbl_bdr_right: PropPair
    tbl_bdr_bottom: PropPair
    tbl_bdr_horz: PropPair
    tbl_bdr_vert: PropPair
    tbl_bdr_dist: PropPair
