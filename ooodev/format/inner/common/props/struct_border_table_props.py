from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class StructBorderTableProps(NamedTuple):
    """Internal Properties"""

    left: PropPair
    top: PropPair
    right: PropPair
    bottom: PropPair
    horz: PropPair
    vert: PropPair
    dist: PropPair
