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
