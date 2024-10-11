from __future__ import annotations
from typing import cast, TYPE_CHECKING
import warnings
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.deployment.package_manager_factory_partial import PackageManagerFactoryPartial

if TYPE_CHECKING:
    from com.sun.star.deployment import thePackageManagerFactory
    from ooodev.loader.inst.lo_inst import LoInst


class ThePackageManagerFactoryComp(ComponentBase, PackageManagerFactoryPartial):
    """
    Class for managing thePackageManagerFactory Component.

    **Deprecated**: This class is deprecated use ``ooodev.adapter.deployment.extension_manager.thePackageManagerFactory`` instead.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: thePackageManagerFactory) -> None:
        """
        Constructor

        Args:
            component (thePackageManagerFactory): UNO Component that implements ``com.sun.star.deployment.thePackageManagerFactory`` service.
        """
        warnings.warn(
            "thePackageManagerFactory is deprecated use ``ooodev.adapter.deployment.extension_manager.thePackageManagerFactory`` instead.",
            DeprecationWarning,
            stacklevel=2,
        )
        ComponentBase.__init__(self, component)
        PackageManagerFactoryPartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> ThePackageManagerFactoryComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ThePackageManagerFactoryComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "com.sun.star.deployment.thePackageManagerFactory"
        if key in lo_inst.cache:
            return cast(ThePackageManagerFactoryComp, lo_inst.cache[key])
        factory = lo_inst.get_singleton("/singletons/com.sun.star.deployment.thePackageManagerFactory")  # type: ignore
        if factory is None:
            raise ValueError("Could not get thePackageManagerFactory singleton.")
        inst = cls(factory)
        lo_inst.cache[key] = inst
        return cast(ThePackageManagerFactoryComp, inst)

    # region Properties
    @property
    def component(self) -> thePackageManagerFactory:
        """thePackageManagerFactory Component"""
        # pylint: disable=no-member
        return cast("thePackageManagerFactory", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
