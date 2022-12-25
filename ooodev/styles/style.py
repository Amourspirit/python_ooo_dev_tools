from __future__ import annotations
from typing import Any

from ..proto.style_obj import StyleObj


class Style:
    @staticmethod
    def apply_style(obj: Any, style: StyleObj) -> None:
        """
        Applies style to object

        Args:
            style (StyleObj): Style object such as ``Font``
            obj (Any): UNO Oject that styles are to be applied.

        Returns:
            None:
        """
        style.apply_style(obj)
