# coding: utf-8
"""
Global Named Events.
"""
from __future__ import annotations
from typing import NamedTuple


class GblNamedEvent(NamedTuple):
    """
    Global Named Events
    """

    PRINTING = "global_printing"
    """Global Printing Event."""

    INPUT_BOX_CREATING = "global_input_box_creating"
    """Global Input Box Creating Event."""

    MSG_BOX_CREATING = "global_msg_box_creating"
    """Global Message Box Creating Event."""

    EVENT_CANCELED = "global_event_canceled"
    """Global Event Canceled Event."""

    RANGE_OBJ_BEFORE_FROM_RANGE = "global_range_obj_before_from_range"

    DOCUMENT_EVENT = "global_document_event"
    """Event is raise when a global document event is triggered."""
