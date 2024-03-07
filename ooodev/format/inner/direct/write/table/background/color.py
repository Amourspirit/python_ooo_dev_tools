# region Imports
from __future__ import annotations
from typing import Tuple
from ooodev.format.inner.direct.calc.background.color import Color as CellColor
from ooodev.format.inner.common.props.cell_background_color_props import CellBackgroundColorProps
from ooodev.utils import color as mColor
from ooodev.utils.color import StandardColor

# endregion Imports


class Color(CellColor):
    """
    Class for Cell Properties Back Color.

    .. seealso::

        - :ref:`help_writer_format_direct_table_background`

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: mColor.Color = StandardColor.AUTO_COLOR) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): Color such as ``CommonColor.LIGHT_BLUE``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_direct_table_background`
        """
        super().__init__(color=color)

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.text.TextTable",
                "com.sun.star.text.CellProperties",
                "com.sun.star.text.TextTableRow",
            )
        return self._supported_services_values

    # endregion overrides

    # region Properties

    @property
    def _props(self) -> CellBackgroundColorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellBackgroundColorProps(
                color="BackColor", is_transparent="BackTransparent"
            )
        return self._props_internal_attributes

    # endregion Properties
