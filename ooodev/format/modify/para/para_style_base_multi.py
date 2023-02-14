from __future__ import annotations
from typing import Tuple

import uno
from ....utils import info as mInfo
from ....utils import lo as mLo
from ...style_base import StyleModifyMulti
from ...writer.style.para.kind import StyleParaKind as StyleParaKind


class ParaStyleBaseMulti(StyleModifyMulti):
    """
    Para Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.ParagraphProperties", "com.sun.star.style.ParagraphStyle")

    def _is_valid_doc(self, obj: object) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)
