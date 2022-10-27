from enum import IntEnum


class SearchMatchKind(IntEnum):
    """Search Match Kind"""

    FULL = 1
    """Full search match. Search criteria must match Exact"""
    FULL_IGNORE_CASE = 2
    """Full search match. Case is ignored. Search criteria must match Exact"""
    PARTIAL = 3
    """Partial match. Search criteria must be in the search result"""
    PARTIAL_IGNORE_CASE = 4
    """Partial match. Case is ignored. Search criteria must be in the search result"""
