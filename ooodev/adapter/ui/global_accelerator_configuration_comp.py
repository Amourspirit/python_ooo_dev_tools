from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.ui import XAcceleratorConfiguration
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.ui.accelerator_configuration_partial import AcceleratorConfigurationPartial
from ooodev.adapter.ui.ui_configuration_events import UIConfigurationEvents

if TYPE_CHECKING:
    from com.sun.star.ui import GlobalAcceleratorConfiguration
    from ooodev.utils.inst.lo.lo_inst import LoInst


class GlobalAcceleratorConfigurationComp(ComponentBase, AcceleratorConfigurationPartial, UIConfigurationEvents):
    """
    Class for managing GlobalAcceleratorConfiguration Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GlobalAcceleratorConfiguration) -> None:
        """
        Constructor

        Args:
            component (GlobalAcceleratorConfiguration): UNO GlobalAcceleratorConfiguration Component that supports ``com.sun.star.ui.GlobalAcceleratorConfiguration`` service.
        """
        # component is a struct
        ComponentBase.__init__(self, component)
        AcceleratorConfigurationPartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        UIConfigurationEvents.__init__(
            self, trigger_args=generic_args, cb=self.__on_ui_configuration_events_add_remove
        )

    # region Lazy Listeners
    def __on_ui_configuration_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addConfigurationListener(self.events_listener_ui_configuration)
        event.remove_callback = True

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by mTextRangePartial.TextRangePartial
        return ("com.sun.star.ui.GlobalAcceleratorConfiguration",)

    # endregion Overrides

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None, *args: Any) -> GlobalAcceleratorConfigurationComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.
            args (Any, optional): One or more args to pass to instance creation.

        Returns:
            ConfigurationProviderComp: The instance with additional classes implemented.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        service = "com.sun.star.ui.GlobalAcceleratorConfiguration"

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XAcceleratorConfiguration, service_name=service, raise_err=True)
        return cls(inst)  # type: ignore

    @property
    def component(self) -> GlobalAcceleratorConfiguration:
        """GlobalAcceleratorConfiguration Component"""
        # pylint: disable=no-member
        return cast("GlobalAcceleratorConfiguration", self._ComponentBase__get_component())  # type: ignore
