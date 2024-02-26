# region Imports
from __future__ import annotations
from typing import Tuple

from ooodev.format.inner.modify.draw.draw_style_base_multi import DrawStyleBaseMulti

# endregion Imports


class LinePropertiesStyleBaseMulti(DrawStyleBaseMulti):
    """
    Line Properties Style Base Multi

    .. versionadded:: 0.17.13
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.LineProperties",
            "com.sun.star.style.Style",
        )
