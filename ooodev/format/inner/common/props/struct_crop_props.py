from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class StructCropProps(NamedTuple):
    """Internal Properties"""

    left: str
    top: str
    right: str
    bottom: str
