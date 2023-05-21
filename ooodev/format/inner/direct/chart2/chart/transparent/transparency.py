from __future__ import annotations
from typing import Any, Tuple

from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as WriteTransparency
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils.data_type.intensity import Intensity


class Transparency(WriteTransparency):
    """
    Chart Fill Transparency.

    .. seealso::

        - :ref:`help_chart2_format_direct_general_transparency`

    .. versionadded:: 0.9.4
    """

    def __init__(self, value: Intensity | int = 0) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.

        Returns:
            None:

        See Also:

            - :ref:`help_chart2_format_direct_general_transparency`
        """
        super().__init__(value=value)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPoint",
                "com.sun.star.chart2.Legend",
            )
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return self._is_obj_service(obj)

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop
