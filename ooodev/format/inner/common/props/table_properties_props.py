from __future__ import annotations
from typing import NamedTuple


class TablePropertiesProps(NamedTuple):
    """Internal Properties"""

    name: str
    width: str
    left: str
    top: str
    right: str
    bottom: str
    is_rel: str
    rel_width: str
    hori_orient: str
