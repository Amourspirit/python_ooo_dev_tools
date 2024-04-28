from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.frame import XModuleManager2
from ooodev.utils.builder.default_builder import DefaultBuilder

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.frame import module_manager2_partial

if TYPE_CHECKING:
    from com.sun.star.frame import ModuleManager  # singleton
    from ooodev.loader.inst.lo_inst import LoInst


class ModuleManagerComp(ComponentProp, module_manager2_partial.ModuleManager2Partial):
    """
    Class for managing ModuleManager Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XModuleManager2) -> None:
        """
        Constructor

        Args:
            component (XModuleManager2): UNO Component that implements ``com.sun.star.frame.ModuleManager`` service.
        """
        ComponentProp.__init__(self, component)
        module_manager2_partial.ModuleManager2Partial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.frame.ModuleManager",)

    # endregion Overrides
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> ModuleManagerComp:
        """
        Creates a new a instance from Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ModuleManager: The new instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XModuleManager2, "com.sun.star.frame.ModuleManager", raise_err=True)
        if inst is None:
            raise ValueError("Could not get ModuleManager singleton.")
        return cls(inst)

    # region Properties
    @property
    def component(self) -> ModuleManager:
        """ModuleManager Component"""
        # pylint: disable=no-member
        return cast("ModuleManager", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """

    return module_manager2_partial.get_builder(component)
