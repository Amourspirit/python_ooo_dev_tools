# region Imports
from __future__ import annotations
from typing import Tuple

from .draw_style_base_multi import DrawStyleBaseMulti

# endregion Imports


class ParaStyleBaseMulti(DrawStyleBaseMulti):
    """
    Paragraph Style Base Multi

    .. versionadded:: 0.17.12
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.style.ParagraphStyle",
        )
