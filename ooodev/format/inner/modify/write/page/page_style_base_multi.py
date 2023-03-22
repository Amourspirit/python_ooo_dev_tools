# region Import
from __future__ import annotations
from typing import Tuple

from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.format.inner.style_base import StyleModifyMulti

# endregion Import


class PageStyleBaseMulti(StyleModifyMulti):
    """
    Page Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")

    def _is_valid_doc(self, obj: object) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)
