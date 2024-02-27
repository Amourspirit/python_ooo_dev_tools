"""
Module for managing the applying of styles.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any

from ooodev.proto import style_obj


class Styler:
    """
    Style methods

    .. versionadded:: 0.9.0
    """

    @staticmethod
    def apply(obj: Any, *styles: style_obj.StyleT, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (Any): UNO Object that styles are to be applied.
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        for style in styles:
            style.apply(obj, **kwargs)
