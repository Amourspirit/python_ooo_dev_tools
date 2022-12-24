from __future__ import annotations
from typing import Any, Tuple, TYPE_CHECKING
import uno
from com.sun.star.beans import XPropertySet
from ..utils import props as mProps

from abc import ABC

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue


class StyleBase(ABC):
    def __init__(self, **kwargs) -> None:
        self._dv = {}
        for (
            key,
            value,
        ) in kwargs.items():
            if not value is None:
                self._dv[key] = value

    def _get(self, key: str) -> Any:
        return self._dv.get(key, None)

    def _set(self, key: str, val: Any) -> bool:
        if val is None:
            if key in self._dv:
                del self._dv[key]
            return False
        self._dv[key] = val
        return True

    def _has(self, key: str) -> bool:
        return key in self._dv

    def apply_style(self, obj: object) -> None:
        if len(self._dv) > 0:
            mProps.Props.set(obj, **self._dv)

    def get_props(self) -> Tuple[PropertyValue, ...]:
        # see: setPropertyValues([in] sequence< com::sun::star::beans::PropertyValue > aProps)
        # https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertyAccess.html#a5ac97dfa6d796f4c794e2350e9130692
        if len(self._dv) == 0:
            return ()
        return mProps.Props.make_props(**self._dv)
