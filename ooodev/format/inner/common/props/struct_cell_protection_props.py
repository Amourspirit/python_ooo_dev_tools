from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class StructCellProtectionProps(NamedTuple):
    """Internal Properties"""

    hide_all: str
    hide_formula: str
    protected: str
    hide_print: str
