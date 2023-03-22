# region Imports
from __future__ import annotations
from typing import Tuple

import uno
from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.format.style_base import StyleModifyMulti
# endregion Imports

class CellStyleBaseMulti(StyleModifyMulti):
    """
    Cell Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle",)
        return self._supported_services_values

    def _is_valid_doc(self, obj: object) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.CALC)
