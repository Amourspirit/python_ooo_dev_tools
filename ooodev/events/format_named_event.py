"""
Write Named Events.
"""

from __future__ import annotations
from typing import NamedTuple


class FormatNamedEvent(NamedTuple):
    """
    Named events for :py:class:`~.format.style_base.StyleBase` class
    """

    STYLE_APPLYING = "format_style_applying"
    """Style applying. Event when Style is applying to an object such as a ``XShape``"""
    STYLE_APPLIED = "format_style_applied"
    """Style Applied. Event when Style has been applied to an object such as a ``XShape``"""

    STYLE_PROPERTY_APPLYING = "format_style_prop_applying"
    """
    Style Property Applying. Event when property is being applied to an object.
    Triggers for each property being set on an UNO Object.
    """
    STYLE_PROPERTY_APPLIED = "format_style_prop_applied"
    """
    Style Property Applied. Event when property has been applied to an object.
    Triggers for each property that has been set on an UNO Object.
    """
    STYLE_PROPERTY_ERROR = "format_style_prop_error"
    """
    Style Property Error. Event when property has failed to be applied for an object.
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
    """Style Setting. This is generally triggered when an a class property is being set."""
    STYLE_SET = "format_style_set"
    """Style Set. This is generally triggered when an a class property has been set."""
    STYLE_MODIFYING = "format_style_modifying"
    """
    Style Modifying. Generally triggered when a style is calling ``_set()``, ``_remove()``, ``_update()``, ``_clear()``
    """
    STYLE_CLEARING = "format_style_clearing"
    """Style Clearing. Generally triggered when style is calling ``_clear()``."""
    STYLE_REMOVING = "format_style_removing"
    """Style Removing. Generally triggered when style is calling ``_remove()``."""
    STYLE_UPDATING = "format_style_updating"
    """Style Updating. Generally triggered when style is calling ``_update``"""
    STYLE_COPYING = "format_style_copying"
    """Style Copying. Generally triggered when style is calling ``copy``"""

    MULTI_STYLE_SETTING = "format_multi_style_setting"
    """Multi Style Setting. Generally triggered when ``StyleMulti._set_style()`` is called"""
    MULTI_STYLE_SET = "format_multi_style_set"
    """Multi Style Set. Generally triggered when ``StyleMulti._set_style()`` is called"""
    MULTI_STYLE_REMOVING = "format_multi_style_removing"
    """Multi Style Removing. Generally triggered when ``StyleMulti._remove_style()`` is called"""
    MULTI_STYLE_REMOVED = "format_multi_style_removed"
    """Multi Style Removed. Generally triggered when ``StyleMulti._remove_style()`` is called"""
    MULTI_STYLE_UPDATING = "format_multi_style_updating"
    """Multi Style Setting. Generally triggered when ``StyleMulti._update_style()`` is called"""
    MULTI_STYLE_UPDATED = "format_multi_style_updated"
    """Multi Style Setting. Generally triggered when ``StyleMulti._update_style()`` is called"""

    STYLE_MULTI_CHILD_APPLYING = "format_multi_child_style_applying"
    """
    Multi Style Child Applying. Event when child style is being applied to an object.
    Triggers for each child style object.
    Generally triggered when ``StyleMulti.apply()`` is called
    """
    STYLE_MULTI_CHILD_APPLIED = "format_multi_child_style_applied"
    """
    Multi Style Child Applied. Event when child style has being applied to an object.
    Triggers for each child style object.
    Generally triggered when ``StyleMulti.apply()`` is called
    """
