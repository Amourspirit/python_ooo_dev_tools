from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class StructSizeProps(NamedTuple):
    """Internal Properties"""

    width: str
    height: str
