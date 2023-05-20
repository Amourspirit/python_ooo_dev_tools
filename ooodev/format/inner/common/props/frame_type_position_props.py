from __future__ import annotations
from typing import NamedTuple


class FrameTypePositionProps(NamedTuple):
    """Internal Properties"""

    hori_orient: str  # HoriOrient
    hori_pos: str  # HoriOrientPosition
    hori_rel: str  # HoriOrientRelation
    vert_orient: str  # VertOrient
    vert_pos: str  # VertOrientPosition
    vert_rel: str  # VertOrientRelation
    txt_flow: str  # IsFollowingTextFlow
    page_toggle: str  # PageToggle
