from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class CellStyleBordersProps(NamedTuple):
    """Internal Properties"""

    shadow: str  # ShadowFormat
    diag_up: str  # DiagonalBLTR2
    diag_dn: str  # DiagonalTLBR2
    bdr_left: str
    bdr_right: str
    bdr_top: str
    bdr_btm: str
    pad_left: str  # ParaLeftMargin
    pad_top: str  # ParaTopMargin
    pad_right: str  # ParaRightMargin
    pad_btm: str  # ParaBottomMargin
