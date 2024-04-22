from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.ui.module_ui_configuration_manager_supplier_partial import (
    ModuleUIConfigurationManagerSupplierPartial,
)

if TYPE_CHECKING:
    from com.sun.star.ui import theModuleUIConfigurationManagerSupplier  # singleton
    from ooodev.loader.inst.lo_inst import LoInst


class TheModuleUIConfigurationManagerSupplierComp(ComponentBase, ModuleUIConfigurationManagerSupplierPartial):
    """
    Class for managing theModuleUIConfigurationManagerSupplier Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: theModuleUIConfigurationManagerSupplier) -> None:
        """
        Constructor

        Args:
            component (theModuleUIConfigurationManagerSupplier): UNO Component that implements ``com.sun.star.ui.theModuleUIConfigurationManagerSupplier`` service.
        """
        ComponentBase.__init__(self, component)
        ModuleUIConfigurationManagerSupplierPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.ui.ModuleUIConfigurationManagerSupplier",)

    # endregion Overrides
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TheModuleUIConfigurationManagerSupplierComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TheModuleUIConfigurationManagerSupplierComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "com.sun.star.ui.theModuleUIConfigurationManagerSupplier"
        if key in lo_inst.cache:
            return cast(TheModuleUIConfigurationManagerSupplierComp, lo_inst.cache[key])

        factory = lo_inst.get_singleton("/singletons/com.sun.star.ui.theModuleUIConfigurationManagerSupplier")  # type: ignore
        if factory is None:
            raise ValueError("Could not get theModuleUIConfigurationManagerSupplier singleton.")
        class_inst = cls(factory)
        lo_inst.cache[key] = class_inst
        return cast(TheModuleUIConfigurationManagerSupplierComp, class_inst)

    # region Properties
    @property
    def component(self) -> theModuleUIConfigurationManagerSupplier:
        """theModuleUIConfigurationManagerSupplier Component"""
        # pylint: disable=no-member
        return cast("theModuleUIConfigurationManagerSupplier", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
