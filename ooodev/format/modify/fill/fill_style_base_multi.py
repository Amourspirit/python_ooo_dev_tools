from __future__ import annotations
from typing import Tuple

import uno
from ....utils import info as mInfo
from ....utils import lo as mLo
from ...kind.format_kind import FormatKind
from ...style_base import StyleModifyMulti


class FillStyleBaseMulti(StyleModifyMulti):
    """
    Fill Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.drawing.FillProperties",)

    def _is_valid_doc(self, obj: object) -> bool:
        valid = mInfo.Info.is_doc_type(obj, mLo.Lo.Service.DRAW)
        if valid:
            return True
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.IMPRESS)
