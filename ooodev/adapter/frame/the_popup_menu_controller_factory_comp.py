from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.frame.ui_controller_factory_partial import UIControllerFactoryPartial


if TYPE_CHECKING:
    from com.sun.star.frame import thePopupMenuControllerFactory  # singleton
    from ooodev.loader.inst.lo_inst import LoInst


class ThePopupMenuControllerFactoryComp(ComponentBase, UIControllerFactoryPartial):
    """
    Class for managing thePopupMenuControllerFactory Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: thePopupMenuControllerFactory) -> None:
        """
        Constructor

        Args:
            component (thePopupMenuControllerFactory): UNO Component that implements ``com.sun.star.frame.thePopupMenuControllerFactory`` service.
        """
        ComponentBase.__init__(self, component)
        UIControllerFactoryPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> ThePopupMenuControllerFactoryComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ThePopupMenuControllerFactoryComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "/singletons/com.sun.star.frame.thePopupMenuControllerFactory"
        if key in lo_inst.cache:
            return cast(ThePopupMenuControllerFactoryComp, lo_inst.cache[key])
        factory = lo_inst.get_singleton(key)  # type: ignore
        if factory is None:
            raise ValueError("Could not get thePopupMenuControllerFactory singleton.")
        result = cls(factory)
        lo_inst.cache[key] = result
        return result

    # region Properties
    @property
    def component(self) -> thePopupMenuControllerFactory:
        """thePopupMenuControllerFactory Component"""
        # pylint: disable=no-member
        return cast("thePopupMenuControllerFactory", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
