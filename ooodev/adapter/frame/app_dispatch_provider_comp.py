from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.frame import XDispatchInformationProvider
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.frame.app_dispatch_provider_partial import AppDispatchProviderPartial

if TYPE_CHECKING:
    from com.sun.star.frame import AppDispatchProvider  # service
    from ooodev.loader.inst.lo_inst import LoInst


class AppDispatchProviderComp(ComponentProp, AppDispatchProviderPartial):
    """
    Class for managing AppDispatchProvider.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDispatchInformationProvider) -> None:
        """
        Constructor

        Args:
            component (XDispatchInformationProvider): UNO Component that implements ``com.sun.star.frame.AppDispatchProvider`` service.
        """
        ComponentProp.__init__(self, component)
        AppDispatchProviderPartial.__init__(self, component=component)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.frame.AppDispatchProvider",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> AppDispatchProviderComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            AppDispatchProviderComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XDispatchInformationProvider, "com.sun.star.frame.AppDispatchProvider", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> AppDispatchProvider:
        """AppDispatchProvider Component"""
        # pylint: disable=no-member
        return cast("AppDispatchProvider", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
