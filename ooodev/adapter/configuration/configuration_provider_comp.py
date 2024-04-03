from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.lang import XMultiServiceFactory
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.lang.multi_service_factory_partial import MultiServiceFactoryPartial

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationProvider  # service
    from ooodev.loader.inst.lo_inst import LoInst


class ConfigurationProviderComp(ComponentBase, MultiServiceFactoryPartial):
    """
    Class for managing table ConfigurationProvider Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ConfigurationProvider) -> None:
        """
        Constructor

        Args:
            component (ConfigurationProvider): UNO ConfigurationProvider Component.
        """
        ComponentBase.__init__(self, component)
        MultiServiceFactoryPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.ConfigurationProvider",)

    # endregion Overrides

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None, *args: Any) -> ConfigurationProviderComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.
            args (Any, optional): One or more args to pass to instance creation.

        Returns:
            ConfigurationProviderComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(
            XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider", args=args, raise_err=True
        )
        return cls(inst)  # type: ignore

    # region Properties
    @property
    def component(self) -> ConfigurationProvider:
        """ConfigurationProvider Component"""
        # pylint: disable=no-member
        return cast("ConfigurationProvider", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
