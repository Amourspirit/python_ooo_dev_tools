# region Imports
from __future__ import annotations
from typing import Tuple

from ooodev.format.inner.modify.draw.draw_style_base_multi import DrawStyleBaseMulti

# endregion Imports


class CharStyleBaseMulti(DrawStyleBaseMulti):
    """
    Char Style Base Multi

    .. versionadded:: 0.17.9
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.CharacterStyle",
        )
