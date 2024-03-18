from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.awt import XUnitConversion

from ooodev.adapter.awt.unit_conversion_partial import UnitConversionPartial
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class UnitConversionComp(ComponentBase, UnitConversionPartial):
    """
    Class for managing Unit Conversion Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, lo_inst: LoInst, component: XUnitConversion | None = None) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst): Lo Instance. This instance is used to create ``component`` is it is not provided.
            component (XUnitConversion, optional): Unit Conversion Component.
                If not provided, it is created from ``lo_inst``.

        Returns:
            None:

        Note:
            ``component`` is automatically created from ``lo_inst`` if it is not provided.
        """
        if component is None:
            window = lo_inst.desktop.get_current_frame().getContainerWindow()
            component = lo_inst.qi(XUnitConversion, window, raise_err=True)

        ComponentBase.__init__(self, component)  # type: ignore
        UnitConversionPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> XUnitConversion:
        """XUnitConversion Component"""
        # pylint: disable=no-member
        return cast(XUnitConversion, self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
