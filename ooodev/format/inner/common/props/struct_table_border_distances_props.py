from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class StructTableBorderDistancesProps(NamedTuple):
    """Internal Properties"""

    left: str
    top: str
    right: str
    bottom: str
    valid_left: str
    valid_top: str
    valid_right: str
    valid_bottom: str
