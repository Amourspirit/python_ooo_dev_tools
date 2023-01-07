"""
Module for managing the applying of styles.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any

from ..proto import style_obj

class Style:
    """
    Style methods

    .. versionadded:: 0.9.0
    """

    @staticmethod
    def apply_style(obj: Any, *styles: style_obj.StyleObj, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (Any): UNO Oject that styles are to be applied.
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        for style in styles:
            style.apply_style(obj, **kwargs)
