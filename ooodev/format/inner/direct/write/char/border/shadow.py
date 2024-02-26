# region Import
from __future__ import annotations
from typing import Any, Tuple, TYPE_CHECKING

import uno
from ooo.dyn.table.shadow_location import ShadowLocation

from ooodev.utils.color import Color
from ooodev.utils.color import StandardColor
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.format.inner.direct.structs.shadow_struct import ShadowStruct

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import


class Shadow(ShadowStruct):
    """
    Shadow struct

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_char_borders`

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

            - :ref:`help_writer_format_direct_char_borders`
        """
        super().__init__(location=location, color=color, transparent=transparent, width=width)

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "CharShadowFormat"
        return self._property_name

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)
