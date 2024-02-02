# region Imports
from __future__ import annotations
from typing import Any, Tuple

from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.format.inner.style_base import StyleModifyMulti

# endregion Imports


class FrameStyleBaseMulti(StyleModifyMulti):
    """
    Frame Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.Style",)
        return self._supported_services_values

    def _is_valid_doc(self, obj: Any) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)
