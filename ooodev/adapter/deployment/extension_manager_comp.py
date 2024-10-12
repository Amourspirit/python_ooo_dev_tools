from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.deployment.extension_manager_partial import ExtensionManagerPartial

if TYPE_CHECKING:
    from com.sun.star.deployment import ExtensionManager
    from ooodev.loader.inst.lo_inst import LoInst


class ExtensionManagerComp(ComponentBase, ExtensionManagerPartial):
    """
    Class for managing ExtensionManager Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ExtensionManager) -> None:
        """
        Constructor

        Args:
            component (ExtensionManager): UNO Component that implements ``com.sun.star.deployment.ExtensionManager`` service.
        """
        ComponentBase.__init__(self, component)
        ExtensionManagerPartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> ExtensionManagerComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ExtensionManagerComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "com.sun.star.deployment.ExtensionManager"
        if key in lo_inst.cache:
            return cast(ExtensionManagerComp, lo_inst.cache[key])
        factory = lo_inst.get_singleton("/singletons/com.sun.star.deployment.ExtensionManager")  # type: ignore
        if factory is None:
            raise ValueError("Could not get ExtensionManager singleton.")
        inst = cls(factory)
        lo_inst.cache[key] = inst
        return cast(ExtensionManagerComp, inst)

    # region Properties
    @property
    def component(self) -> ExtensionManager:
        """ExtensionManager Component"""
        # pylint: disable=no-member
        return cast("ExtensionManager", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
