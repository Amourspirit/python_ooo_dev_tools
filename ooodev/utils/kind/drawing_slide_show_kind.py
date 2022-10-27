from enum import IntEnum


class DrawingSlideShowKind(IntEnum):
    """DrawPage slide show change constants"""

    AUTO_CHANGE = 1
    """Everything (page change, animation effects) is automatic"""
    CLICK_ALL_CHANGE = 0
    """A mouse-click triggers the next animation effect or page change"""
    CLICK_PAGE_CHANGE = 2
    """Animation effects run automatically, but the user must click on the page to change it"""
