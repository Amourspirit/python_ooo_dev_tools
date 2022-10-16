# coding: utf-8
"""
Draw Named Events.
"""
from __future__ import annotations
from typing import NamedTuple


class DrawNamedEvent(NamedTuple):
    """
    Named events for utils.draw.Draw class
    """

    GET_SHAPES_ERROR = "draw_get_shapes_error"
    """Draw get_shapes error command see :py:meth:`Draw.get_shapes() <.utils.draw.Draw.get_shapes>`"""