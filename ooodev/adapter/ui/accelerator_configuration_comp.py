from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.ui import accelerator_configuration_partial
from ooodev.adapter.ui.ui_configuration_events import UIConfigurationEvents
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.adapter.form.reset_partial import ResetPartial
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.adapter.util.changes_notifier_partial import ChangesNotifierPartial
from ooodev.adapter.util.changes_events import ChangesEvents

if TYPE_CHECKING:
    from com.sun.star.ui import XAcceleratorConfiguration


class _AcceleratorConfigurationComp(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component


class AcceleratorConfigurationComp(
    _AcceleratorConfigurationComp,
    accelerator_configuration_partial.AcceleratorConfigurationPartial,
    ResetPartial,
    UIConfigurationEvents,
    ResetEvents,
    ChangesNotifierPartial,
    ChangesEvents,
    CompDefaultsPartial,
):
    """
    Class for managing XAcceleratorConfiguration Component.
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
            name="ooodev.adapter.ui.accelerator_configuration_comp.AcceleratorConfigurationComp",
            base_class=_AcceleratorConfigurationComp,
        )
        return inst

    def __init__(self, component: XAcceleratorConfiguration) -> None:
        """
        Constructor

        Args:
            component (XAcceleratorConfiguration): UNO Component that implements ``com.sun.star.ui.XAcceleratorConfiguration`` interface.
        """
        # this it not actually called as __new__ is overridden
        pass

    @property
    def component(self) -> XAcceleratorConfiguration:
        """XAcceleratorConfiguration Component"""
        # pylint: disable=no-member
        return cast("XAcceleratorConfiguration", self._ComponentBase__get_component())  # type: ignore


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
    builder.auto_add_interface("com.sun.star.lang.XTypeProvider")
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
