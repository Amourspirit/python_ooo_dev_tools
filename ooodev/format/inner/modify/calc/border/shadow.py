from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.table.shadow_location import ShadowLocation

from ooodev.utils.color import Color, StandardColor
from ooodev.format.inner.direct.calc.border.shadow import Shadow as DirectShadow

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class Shadow(DirectShadow):
    """
    Border Shadow

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float | UnitT = 1.76,
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow.
                Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (:py:data:`~.utils.color.Color`, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, UnitT, optional): contains the size of the shadow (in ``mm`` units)
                or :ref:`proto_unit_obj`. Defaults to ``1.76``.

        Raises:
            ValueError: If ``color`` or ``width`` are less than zero.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_borders`
        """
        super().__init__(location=location, color=color, transparent=transparent, width=width)
