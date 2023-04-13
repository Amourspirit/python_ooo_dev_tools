from __future__ import annotations
from typing import Tuple

from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as WriteTransparency
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils import lo as mLo


class Transparency(WriteTransparency):
    """
    Chart Wall Fill Transparency.

    .. versionadded:: 0.9.4
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: object) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop
