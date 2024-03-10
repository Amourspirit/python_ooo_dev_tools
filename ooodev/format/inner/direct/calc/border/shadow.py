from __future__ import annotations
from typing import Any, Tuple
import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.units.unit_obj import UnitT
from ooodev.utils.color import Color, StandardColor
from ooodev.format.inner.direct.structs.shadow_struct import ShadowStruct


class Shadow(ShadowStruct):
    """
    Shadow struct

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_borders`
        - :ref:`help_writer_format_direct_table_borders`

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

        Hint:
            - ``ShadowLocation`` can be imported from ``ooo.dyn.table.shadow_location``

        See Also:
            - :ref:`help_calc_format_direct_cell_borders`
            - :ref:`help_writer_format_direct_table_borders`
        """
        super().__init__(location=location, color=color, transparent=transparent, width=width)

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle", "com.sun.star.table.CellProperties")
        return self._supported_services_values
