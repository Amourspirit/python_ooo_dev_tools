from __future__ import annotations
from typing import NamedTuple


class FrameTypeSizeProps(NamedTuple):
    """Internal Properties"""

    width: str  # Width
    height: str  # Height
    rel_width: str  # RealitiveWidth ( int )
    rel_height: str  # RealitiveHeight ( int )
    rel_width_relation: str  # RelativeWidthRelation
    rel_height_relation: str  # RelativeHeighRelation
    size_type: str  # SizeType
    width_type: str  # WidthType
