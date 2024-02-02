# region Imports
from __future__ import annotations
from typing import Any, Tuple

from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.format.inner.style_base import StyleModifyMulti

# endregion Imports


class DrawStyleBaseMulti(StyleModifyMulti):
    """
    Base class for Draw Multi Styles

    .. versionadded:: 0.17.13
    """

    def _supported_services(self) -> Tuple[str, ...]:
        raise NotImplementedError

    def _is_valid_doc(self, obj: Any) -> bool:
        return mInfo.Info.support_service(obj, str(mLo.Lo.Service.DRAW), str(mLo.Lo.Service.IMPRESS))
