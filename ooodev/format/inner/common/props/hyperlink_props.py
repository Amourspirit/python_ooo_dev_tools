from __future__ import annotations
from typing import NamedTuple


class HyperlinkProps(NamedTuple):
    name: str  # HyperLinkName
    target: str  # HyperLinkTarget
    url: str  # HyperLinkURL
    visited: str  # VisitedCharStyleName
    unvisited: str  # UnvisitedCharStyleName
