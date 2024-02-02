from __future__ import annotations
from typing import Any, Tuple

from ooodev.format.inner.direct.write.fill.area.fill_color import FillColor
from ooodev.utils import color as mColor
from ooodev.loader import lo as mLo


class Color(FillColor):
    """
    Class for Chart Wall/Floor Fill Color.

    .. seealso::

        - :ref:`help_chart2_format_direct_wall_floor_area`

    .. versionadded:: 0.9.4
    """

    def __init__(self, color: mColor.Color = mColor.Color(-1)) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_wall_floor_area`
        """
        super().__init__(color=color)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
