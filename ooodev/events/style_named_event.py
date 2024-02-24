"""
Write Named Events.
"""

from __future__ import annotations
from typing import NamedTuple


class StyleNameEvent(NamedTuple):
    """
    Named events for Styles.
    """

    STYLE_APPLYING = "styling_style_applying"
    """Style applying. Event when Style is applying to an object."""
    STYLE_APPLIED = "styling_style_applied"
    """Style Applied. Event when Style has been applied to an object."""

    STYLE_BACKING_UP = "styling_style_backing_up"
    """Style Backing up"""
    STYLE_BACKED_UP = "styling_style_backed_up"
    """Style Backed up"""
    STYLE_RESTORING = "styling_style_restoring"
    """Style Restoring up"""
    STYLE_RESTORED = "styling_style_restored"
    """Style Restored up"""

    STYLE_NAME_APPLYING = "styling_style_name_applying"
    """Style Name Applying"""
    STYLE_NAME_APPLIED = "styling_style_name_applied"
    """Style Name Applied"""
