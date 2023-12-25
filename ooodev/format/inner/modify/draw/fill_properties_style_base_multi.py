# region Imports
from __future__ import annotations
from typing import Any, Tuple

from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.format.inner.style_base import StyleModifyMulti

# endregion Imports


class FillPropertiesStyleBaseMulti(StyleModifyMulti):
    """
    Fill Properties Style Base Multi

    .. versionadded:: 0.17.9
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.style.Style",
        )

    def _is_valid_doc(self, obj: Any) -> bool:
        return mInfo.Info.support_service(obj, str(mLo.Lo.Service.DRAW), str(mLo.Lo.Service.IMPRESS))
