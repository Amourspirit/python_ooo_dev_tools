# region Imports
from __future__ import annotations
from typing import Any, Tuple

from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.format.inner.style_base import StyleModifyMulti

# endregion Imports


class CharStyleBaseMulti(StyleModifyMulti):
    """
    Char Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.CharacterStyle",
            "com.sun.star.style.ParagraphStyle",
        )

    def _is_valid_doc(self, obj: Any) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)
