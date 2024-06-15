from __future__ import annotations
from typing import Any, cast, overload, TYPE_CHECKING
import uno

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder

from ooodev.adapter.sheet import spreadsheet_view_comp
from ooodev.adapter.sheet import spreadsheet_view_settings_comp
from ooodev.calc import calc_cell as mCalcCell
from ooodev.calc import calc_cell_cursor as mCalcCellCursor
from ooodev.calc import calc_cell_range as mCalcCellRange
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils import info as mInfo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.adapter.ui.context_menu_interception_partial import ContextMenuInterceptionPartial
from ooodev.adapter.ui.context_menu_interceptor_events import ContextMenuInterceptorEvents
from ooodev.events.args.generic_args import GenericArgs


from ooodev.adapter.awt.user_input_interception_partial import UserInputInterceptionPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.datatransfer.transferable_supplier_partial import TransferableSupplierPartial
from ooodev.adapter.frame.controller_border_partial import ControllerBorderPartial
from ooodev.adapter.frame.controller2_partial import Controller2Partial
from ooodev.adapter.frame.dispatch_information_provider_partial import DispatchInformationProviderPartial
from ooodev.adapter.frame.dispatch_provider_partial import DispatchProviderPartial
from ooodev.adapter.frame.infobar_provider_partial import InfobarProviderPartial
from ooodev.adapter.frame.title_change_broadcaster_partial import TitleChangeBroadcasterPartial
from ooodev.adapter.frame.title_partial import TitlePartial
from ooodev.adapter.lang.initialization_partial import InitializationPartial
from ooodev.adapter.lang.type_provider_partial import TypeProviderPartial
from ooodev.adapter.lang.uno_tunnel_partial import UnoTunnelPartial
from ooodev.adapter.task.status_indicator_supplier_partial import StatusIndicatorSupplierPartial
from ooodev.adapter.frame.border_resize_events import BorderResizeEvents
from ooodev.adapter.frame.title_change_events import TitleChangeEvents


if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetView

    # from com.sun.star.sheet import SpreadsheetView  # service
    # from com.sun.star.sheet import SpreadsheetViewSettings  # service
    from ooodev.calc.calc_doc import CalcDoc


class _CalcSheetView(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetView",)

    # endregion Overrides

    @overload
    def select(self, selection: mCalcCellRange.CalcCellRange) -> bool:
        """
        Selects the cells represented by the selection.

        Args:
            selection (CalcCellRange): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        ...

    @overload
    def select(self, selection: mCalcCell.CalcCell) -> bool:
        """
        Selects the cells represented by the selection.

        Args:
            selection (CalcCell): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        ...

    @overload
    def select(self, selection: mCalcCellCursor.CalcCellCursor) -> bool:
        """
        Selects the cells represented by the selection.

        Args:
            selection (CalcCellCursor): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        ...

    def select(self, selection: Any) -> bool:
        """
        Selects the object represented by xSelection if it is known and selectable in this object.

        Args:
            selection (Any): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        if mInfo.Info.is_instance(selection, mCalcCellRange.CalcCellRange):
            cursor = selection.create_cursor()
            return self.component.select(cursor.component)
        elif mInfo.Info.is_instance(selection, mCalcCell.CalcCell):
            cursor = selection.create_cursor()
            return self.component.select(cursor.component)
        elif mInfo.Info.is_instance(selection, mCalcCellCursor.CalcCellCursor):
            cell_rng = selection.get_calc_cell_range()
            cursor = cell_rng.create_cursor()
            return self.component.select(cursor.component)

        return self.component.select(selection)

    def get_selection(self) -> Any:
        """
        Returns the current selection.

        Returns:
            Any: Selection
        """
        return self.component.getSelection()

    # region Properties
    @property
    def __class__(self):
        # pretend to be a CalcSheetView class
        return CalcSheetView

    # endregion Properties


class CalcSheetView(
    _CalcSheetView,
    LoInstPropsPartial,
    spreadsheet_view_comp.SpreadsheetViewComp,
    spreadsheet_view_settings_comp.SpreadsheetViewSettingsComp,
    ContextMenuInterceptorEvents,
    ContextMenuInterceptionPartial,
    PropPartial,
    StylePartial,
    ServicePartial,
    TheDictionaryPartial,
    CalcDocPropPartial,
    CompDefaultsPartial,
    UserInputInterceptionPartial,
    PropertySetPartial,
    TransferableSupplierPartial,
    ControllerBorderPartial,
    Controller2Partial,
    DispatchInformationProviderPartial,
    DispatchProviderPartial,
    InfobarProviderPartial,
    TitleChangeBroadcasterPartial,
    TitlePartial,
    InitializationPartial,
    TypeProviderPartial,
    UnoTunnelPartial,
    StatusIndicatorSupplierPartial,
    BorderResizeEvents,
    TitleChangeEvents,
):
    # pylint: disable=unused-argument
    def __new__(cls, owner: CalcDoc, view: XSpreadsheetView, *args, **kwargs):
        builder = get_builder(component=view)
        builder_helper.builder_add_comp_defaults(builder)
        builder_helper.builder_add_the_dictionary_partial(builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)

        clz = builder.get_class_type(
            name="ooodev.calc.calc_sheet_view.CalcSheetView",
            base_class=_CalcSheetView,
        )
        inst = clz(builder.component)
        lo_inst = kwargs.get("lo_inst", mLo.Lo.current_lo)
        if builder.has_type(LoInstPropsPartial):
            builder.pop_type(LoInstPropsPartial)
            LoInstPropsPartial.__init__(inst, lo_inst=lo_inst)
        if builder.has_type(PropPartial):
            builder.pop_type(PropPartial)
            PropPartial.__init__(inst, component=builder.component, lo_inst=lo_inst)
        if builder.has_type(StylePartial):
            builder.pop_type(StylePartial)
            StylePartial.__init__(inst, component=builder.component)
        if builder.has_type(ServicePartial):
            builder.pop_type(ServicePartial)
            ServicePartial.__init__(inst, component=builder.component, lo_inst=lo_inst)

        if builder.has_type(CalcDocPropPartial):
            builder.pop_type(CalcDocPropPartial)
            CalcDocPropPartial.__init__(inst, obj=owner)
        if builder.has_type(ContextMenuInterceptorEvents):
            builder.pop_type(ContextMenuInterceptorEvents)
            ContextMenuInterceptorEvents.__init__(inst, component=view, trigger_args=GenericArgs(view=inst))  # type: ignore
        # init any remaining types
        builder.init_classes(inst)
        return inst

    def __init__(self, owner: CalcDoc, view: XSpreadsheetView, lo_inst: LoInst | None = None) -> None:
        pass


def get_builder(component: Any, **kwargs: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    auto_interface = kwargs.pop("auto_interface", True)

    builder.add_import(
        name="ooodev.utils.partial.lo_inst_props_partial.LoInstPropsPartial",
    )

    if auto_interface:
        builder.auto_interface()

    builder.merge(spreadsheet_view_comp.get_builder(component, auto_interface=False))
    builder.merge(spreadsheet_view_settings_comp.get_builder(component, auto_interface=False))

    builder.add_import(name="ooodev.utils.partial.prop_partial.PropPartial")
    builder.add_import(name="ooodev.format.inner.style_partial.StylePartial")
    builder.add_import(name="ooodev.utils.partial.service_partial.ServicePartial")
    builder.add_import(name="ooodev.calc.partial.calc_doc_prop_partial.CalcDocPropPartial")

    # ContextMenuInterceptorEvents needs to be added as a regular import It does not auto attach a listener.
    # handled in __new__()
    builder.add_import(name="ooodev.adapter.ui.context_menu_interceptor_events.ContextMenuInterceptorEvents")

    # builder.add_event(
    #     module_name="ooodev.adapter.ui.context_menu_interceptor_events",
    #     class_name="ContextMenuInterceptorEvents",
    #     uno_name="com.sun.star.ui.XContextMenuInterception",
    #     optional=True,
    # )

    builder.add_event(
        module_name="ooodev.adapter.view.selection_change_events",
        class_name="SelectionChangeEvents",
        uno_name="com.sun.star.view.XSelectionSupplier",
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.frame.border_resize_events",
        class_name="BorderResizeEvents",
        uno_name="com.sun.star.frame.XControllerBorder",
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.frame.title_change_events",
        class_name="TitleChangeEvents",
        uno_name="com.sun.star.frame.XTitleChangeBroadcaster",
        optional=True,
    )
    builder_helper.builder_add_property_change_implement(builder)
    builder_helper.builder_add_property_veto_implement(builder)

    return builder
