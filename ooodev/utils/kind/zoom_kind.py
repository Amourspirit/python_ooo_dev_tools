from __future__ import annotations
from enum import IntEnum


class ZoomKind(IntEnum):
    OPTIMAL = 0
    """
    The page content width (excluding margins) at the current selection is fit into the view.
    """
    PAGE_WIDTH = 1
    """
    The page width at the current selection is fit into the view.
    """
    ENTIRE_PAGE = 2
    """
    A complete page of the document is fit into the view.
    """
    BY_VALUE = 3
    """
    The zoom is relative and is to be set via the property ViewSettings.ZoomValue.
    """
    PAGE_WIDTH_EXACT = 4
    """
    The page width at the current selection is fit into the view, with the view ends exactly at the end of the page.
    """
    ZOOM_50_PERCENT = 1000
    """Zoom 50%"""
    ZOOM_75_PERCENT = 1001
    """Zoom 75%"""
    ZOOM_100_PERCENT = 1002
    """Zoom 100%"""
    ZOOM_150_PERCENT = 1003
    """Zoom 150%"""
    ZOOM_200_PERCENT = 1004
    """Zoom 200%"""
