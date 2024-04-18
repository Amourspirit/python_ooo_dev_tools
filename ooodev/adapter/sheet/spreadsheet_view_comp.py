from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.utils.builder.init_kind import InitKind
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.adapter.awt.enhanced_mouse_click_events import EnhancedMouseClickEvents
from ooodev.adapter.awt.key_events import KeyEvents
from ooodev.adapter.awt.mouse_click_events import MouseClickEvents
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.adapter.view.selection_supplier_partial import SelectionSupplierPartial
from ooodev.adapter.sheet.activation_broadcaster_partial import ActivationBroadcasterPartial
from ooodev.adapter.sheet.activation_event_events import ActivationEventEvents
from ooodev.adapter.sheet.enhanced_mouse_click_broadcaster_partial import EnhancedMouseClickBroadcasterPartial
from ooodev.adapter.sheet.range_selection_change_events import RangeSelectionChangeEvents
from ooodev.adapter.sheet.range_selection_partial import RangeSelectionPartial
from ooodev.adapter.sheet.spreadsheet_view_partial import SpreadsheetViewPartial
from ooodev.adapter.sheet.view_freezable_partial import ViewFreezablePartial
from ooodev.adapter.sheet.view_splitable_partial import ViewSplitablePartial
from ooodev.adapter.view.form_layer_access_partial import FormLayerAccessPartial
from ooodev.adapter.sheet import spreadsheet_view_pane_comp
from ooodev.adapter.awt.key_handler_events import KeyHandlerEvents

if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetView  # service
    from com.sun.star.sheet import SpreadsheetViewPane  # service


class _SpreadsheetViewComp(spreadsheet_view_pane_comp._SpreadsheetViewPaneComp):

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetView",)

    # endregion Overrides

    # region Properties
    @property
    def __class__(self):
        # pretend to be a SpreadsheetViewComp class
        return SpreadsheetViewComp

    # endregion Properties


class SpreadsheetViewComp(
    _SpreadsheetViewComp,
    ActivationBroadcasterPartial,
    ActivationEventEvents,
    EnhancedMouseClickBroadcasterPartial,
    IndexAccessPartial["SpreadsheetViewPane"],
    EnumerationAccessPartial["SpreadsheetViewPane"],
    FormLayerAccessPartial,
    RangeSelectionPartial,
    SelectionSupplierPartial,
    SpreadsheetViewPartial,
    ViewFreezablePartial,
    ViewSplitablePartial,
    EnhancedMouseClickEvents,
    KeyHandlerEvents,
    MouseClickEvents,
    RangeSelectionChangeEvents,
    SelectionChangeEvents,
    PropertyChangeImplement,
    VetoableChangeImplement,
    CompDefaultsPartial,
):
    """
    Class for managing Spreadsheet View Component.
    """

    # pylint: disable=unused-argument
    def __new__(cls, component: Any, *args, **kwargs):
        builder = cast(
            DefaultBuilder,
            spreadsheet_view_pane_comp.SpreadsheetViewPaneComp.__new__(
                cls, component, _builder_only=True, *args, **kwargs
            ),
        )

        local_builder = get_builder(component=component, _for_new=True)
        builder.merge(local_builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.adapter.sheet.spreadsheet_view_comp.SpreadsheetViewComp",
            base_class=_SpreadsheetViewComp,
        )
        return inst

    def __init__(self, component: SpreadsheetView) -> None:
        """
        Constructor

        Args:
            component (SpreadsheetView): UNO Spreadsheet View Component
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> SpreadsheetView:
        """Spreadsheet View Component"""
        # pylint: disable=no-member
        return cast("SpreadsheetView", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, **kwargs: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    for_new = kwargs.get("_for_new", False)
    auto_interface = kwargs.get("auto_interface", True)
    if for_new:
        builder = DefaultBuilder(component)
        if auto_interface:
            builder.auto_interface()
    else:
        builder = spreadsheet_view_pane_comp.get_builder(component=component, **kwargs)

    builder.add_event(
        module_name="ooodev.adapter.awt.enhanced_mouse_click_events",
        class_name="EnhancedMouseClickEvents",
        uno_name="com.sun.star.sheet.XEnhancedMouseClickBroadcaster",
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.awt.key_handler_events",
        class_name="KeyHandlerEvents",
        uno_name=("com.sun.star.awt.XUserInputInterception", "com.sun.star.awt.XExtendedToolkit"),
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.awt.mouse_click_events",
        class_name="MouseClickEvents",
        uno_name="com.sun.star.awt.XUserInputInterception",
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.sheet.range_selection_change_events",
        class_name="RangeSelectionChangeEvents",
        uno_name="com.sun.star.sheet.XRangeSelection",
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.view.selection_change_events",
        class_name="SelectionChangeEvents",
        uno_name="com.sun.star.view.XSelectionSupplier",
        optional=True,
    )
    builder.add_import(
        name="ooodev.adapter.beans.property_change_implement.PropertyChangeImplement",
        uno_name="com.sun.star.beans.XPropertySet",
        optional=True,
        init_kind=InitKind.COMPONENT,
        check_kind=CheckKind.INTERFACE,
    )
    builder.add_import(
        name="ooodev.adapter.beans.vetoable_change_implement.VetoableChangeImplement",
        uno_name="com.sun.star.beans.XPropertySet",
        optional=True,
        init_kind=InitKind.COMPONENT,
        check_kind=CheckKind.INTERFACE,
    )
    return builder
