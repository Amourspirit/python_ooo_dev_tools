# region Imports
from __future__ import annotations
from typing import Any, Tuple

from ooodev.format.inner.style_base import StyleModifyMulti
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo

# endregion Imports


class FillStyleBaseMulti(StyleModifyMulti):
    """
    Fill Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.drawing.FillProperties",)

    def _is_valid_doc(self, obj: Any) -> bool:
        return mInfo.Info.support_service(obj, str(mLo.Lo.Service.DRAW), str(mLo.Lo.Service.IMPRESS))
