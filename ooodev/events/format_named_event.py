"""
Write Named Events.
"""
from __future__ import annotations


class FormatNamedEvent:
    """
    Named events for :py:class:`~.format.style_base.StyleBase` class
    """

    STYLE_APPLYING = "format_style_applying"
    """Style applying. Event when Style is being applied to an object such as a ``XShape``"""
    STYLE_APPLIED = "format_style_applyed"
    """Style Applied. Event when Style has been applied to an object such as a ``XShape``"""

    STYLE_PROPERTY_APPLYING = "format_style_prop_applying"
    """
    Style Property Applying. Event when property is being applied to an object.
    Triggers for each property being set on an UNO Object.
    """
    STYLE_PROPERTY_APPLIED = "format_style_prop_applied"
    """
    Style Property Applying. Event when property has been applied to an object.
    Triggers for each property that has been set on an UNO Object.
    """
    STYLE_BACKING_UP = "format_style_backing_up"
    """Style Backing up"""
    STYLE_BACKED_UP = "format_style_backed_up"
    """Style Backed up"""
    STYLE_PROPERTY_RESTORING = "format_style_prop_restoring"
    """Style Restoring up"""
    STYLE_PROPERTY_RESTORED = "format_style_prop_restored"
    """Style Restored up"""
    STYLE_SETTING = "format_style_setting"
    """Style Setting. This is generally when an a class property is being set."""
    STYLE_SET = "format_style_set"
    """Style Set. This is generally when an a class property has been set."""
    STYLE_MODIFING = "format_style_modifing"
    """Style Modifing. Generally when a style is calling ``_set()``, ``_remove()``, ``_update()``, ``_clear()``"""
    STYLE_CLEARING = "format_style_clearing"
    """Style Clearing. Generally when style is calling ``_clear()``."""
    STYLE_REMOVING = "format_style_removing"
    """Style Removing. Generally when style is calling ``_remove()``."""
    STYLE_UPDATING = "format_style_updating"
    """Style Updating. Generall when style is calling ``_update``"""
