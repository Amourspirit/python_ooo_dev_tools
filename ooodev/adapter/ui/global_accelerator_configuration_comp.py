from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.ui import XAcceleratorConfiguration

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.ui import accelerator_configuration_partial
from ooodev.adapter.ui.ui_configuration_events import UIConfigurationEvents
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.ui import GlobalAcceleratorConfiguration
    from ooodev.utils.inst.lo.lo_inst import LoInst


class _GlobalAcceleratorConfigurationComp(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.ui.GlobalAcceleratorConfiguration",)


class GlobalAcceleratorConfigurationComp(
    _GlobalAcceleratorConfigurationComp,
    accelerator_configuration_partial.AcceleratorConfigurationPartial,
    UIConfigurationEvents,
    CompDefaultsPartial,
):
    """
    Class for managing GlobalAcceleratorConfiguration Component.
    """

    # pylint: disable=unused-argument
    def __new__(cls, component: Any, *args, **kwargs):
        builder = get_builder(component=component)
        builder_helper.builder_add_comp_defaults(builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.adapter.ui.global_accelerator_configuration_comp.GlobalAcceleratorConfigurationComp",
            base_class=_GlobalAcceleratorConfigurationComp,
        )
        return inst

    def __init__(self, component: GlobalAcceleratorConfiguration) -> None:
        """
        Constructor

        Args:
            component (GlobalAcceleratorConfiguration): UNO GlobalAcceleratorConfiguration Component that supports ``com.sun.star.ui.GlobalAcceleratorConfiguration`` service.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Class Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> GlobalAcceleratorConfigurationComp:
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

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "com.sun.star.ui.GlobalAcceleratorConfiguration"
        if key in lo_inst.cache:
            return cast(GlobalAcceleratorConfigurationComp, lo_inst.cache[key])
        service = "com.sun.star.ui.GlobalAcceleratorConfiguration"
        inst = lo_inst.create_instance_mcf(XAcceleratorConfiguration, service_name=service, raise_err=True)
        class_inst = cls(inst)  # type: ignore
        lo_inst.cache[key] = class_inst
        return cast(GlobalAcceleratorConfigurationComp, class_inst)

    # endregion Class Methods

    # region Properties
    @property
    def component(self) -> GlobalAcceleratorConfiguration:
        """GlobalAcceleratorConfiguration Component"""
        # pylint: disable=no-member
        return cast("GlobalAcceleratorConfiguration", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.merge(accelerator_configuration_partial.get_builder(component))
    builder.auto_add_interface("com.sun.star.form.XReset")
    builder.auto_add_interface("com.sun.star.util.XChangesNotifier")

    builder.add_event(
        module_name="ooodev.adapter.ui.ui_configuration_events",
        class_name="UIConfigurationEvents",
        uno_name="com.sun.star.ui.XUIConfiguration",
        optional=True,
        check_kind=CheckKind.INTERFACE,
    )
    builder.add_event(
        module_name="ooodev.adapter.form.reset_events",
        class_name="ResetEvents",
        uno_name="com.sun.star.form.XReset",
        optional=True,
        check_kind=CheckKind.INTERFACE,
    )
    builder.add_event(
        module_name="ooodev.adapter.util.changes_events",
        class_name="ChangesEvents",
        uno_name="com.sun.star.util.XChangesNotifier",
        optional=True,
        check_kind=CheckKind.INTERFACE,
    )
    return builder
