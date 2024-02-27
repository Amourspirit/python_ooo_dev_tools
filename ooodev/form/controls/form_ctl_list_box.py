from __future__ import annotations
from typing import Any, cast, Iterable, Tuple, TYPE_CHECKING
import contextlib
import uno

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.adapter.awt.item_list_events import ItemListEvents
from ooodev.adapter.form.change_events import ChangeEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.form_component_kind import FormComponentKind

from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import ListBox as ControlModel  # service
    from com.sun.star.form.control import ListBox as ControlView  # service
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlListBox(FormCtlBase, ActionEvents, ChangeEvents, ItemEvents, ItemListEvents, ResetEvents):
    """``com.sun.star.form.component.ListBox`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.ListBox`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            None:

        Note:
            If the :ref:`LoContext <ooodev.utils.context.lo_context.LoContext>` manager is use before this class is instantiated,
            then the Lo instance will be set using the current Lo instance. That the context manager has set.
            Generally speaking this means that there is no need to set ``lo_inst`` when instantiating this class.

        See Also:
            :ref:`ooodev.form.Forms`.
        """
        FormCtlBase.__init__(self, ctl=ctl, lo_inst=lo_inst)
        generic_args = self._get_generic_args()
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)
        ChangeEvents.__init__(self, trigger_args=generic_args, cb=self._on_change_add_remove)
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_add_remove)
        ItemListEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_list_add_remove)
        ResetEvents.__init__(self, trigger_args=generic_args, cb=self._on_reset_add_remove)

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addActionListener(self.events_listener_action)
        event.remove_callback = True

    def _on_change_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addChangeListener(self.events_listener_change)
        event.remove_callback = True

    def _on_item_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addItemListener(self.events_listener_item)
        event.remove_callback = True

    def _on_reset_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addResetListener(self.events_listener_reset)
        event.remove_callback = True

    def _on_item_list_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addItemListListener(self.events_listener_item_list)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    if TYPE_CHECKING:
        # override the methods to provide type hinting
        def get_view(self) -> ControlView:
            """Gets the view of this control"""
            return cast("ControlView", super().get_view())

        def get_model(self) -> ControlModel:
            """Gets the model for this control"""
            return cast("ControlModel", super().get_model())

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.LIST_BOX

    # endregion Overrides

    # region Methods

    def set_list_data(self, data: Iterable[str]) -> None:
        """
        Sets the list data

        Args:
            data (Iterable[str]): List box entries
        """
        if not data:
            return
        ctl_props = self.get_property_set()
        if ctl_props is None:
            return
        uno_strings = uno.Any("[]string", tuple(data))  # type: ignore
        uno.invoke(ctl_props, "setPropertyValue", ("StringItemList", uno_strings))  # type: ignore

    def get_item(self, index: int) -> str:
        """
        Gets the item at the specified index

        Args:
            index (int): Index of the item to get

        Returns:
            The item at the specified index
        """
        return self.view.getItem(index)

    # endregion Methods

    # region Properties
    @property
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.Border = value.value

    @property
    def drop_down(self) -> bool:
        """Gets/Sets the DropDown property"""
        return self.model.Dropdown

    @drop_down.setter
    def drop_down(self, value: bool) -> None:
        """Sets the DropDown property"""
        self.model.Dropdown = value

    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

    @property
    def help_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def help_url(self) -> str:
        """Gets/Sets the help url"""
        return self.model.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.model.HelpURL = value

    @property
    def line_count(self) -> int:
        """Gets/Sets the line count"""
        return self.model.LineCount

    @line_count.setter
    def line_count(self, value: int) -> None:
        """Sets the line count"""
        self.model.LineCount = value

    @property
    def list_count(self) -> int:
        """Gets the number of items in the list"""
        with contextlib.suppress(Exception):
            return self.view.getItemCount()
        return 0

    @property
    def list_index(self) -> int:
        """
        Gets which item is selected

        Returns:
            Index of the first selected item or ``-1`` if no items are selected.
        """
        with contextlib.suppress(Exception):
            selected_items = self.selected_items
            if len(selected_items) > 0:
                return selected_items[0]
        return -1

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def multi_selection(self) -> bool:
        """Gets/Sets the MultiSelection property"""
        return self.model.MultiSelection

    @multi_selection.setter
    def multi_selection(self, value: bool) -> None:
        """Sets the MultiSelection property"""
        self.model.MultiSelection = value

    @property
    def printable(self) -> bool:
        """Gets/Sets the printable property"""
        return self.model.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.model.Printable = value

    @property
    def read_only(self) -> bool:
        """Gets/Sets the read only property"""
        return self.model.ReadOnly

    @read_only.setter
    def read_only(self, value: bool) -> None:
        self.model.ReadOnly = value

    @property
    def row_source(self) -> Tuple[str, ...]:
        """
        Gets/Sets the row source.

        When setting the row source, the list box will be cleared and the new items will be added.

        The value passed in can be any iterable string such as a list or tuple of strings.
        """
        with contextlib.suppress(Exception):
            return self.model.StringItemList
        return ()

    @row_source.setter
    def row_source(self, value: Iterable[str]) -> None:
        """Sets the row source"""
        self.set_list_data(value)

    @property
    def selected_items(self) -> Tuple[int, ...]:
        """
        Gets/Sets the selected items.

        When setting the selected items, any iterable of integers can be passed in
        such as a list or tuple of integers.

        Returns:
            Tuple[int, ...]: Tuple of selected items

        .. versionchanged:: 0.19.1
            Allows setting the selected items
        """
        return cast(Tuple[int, ...], self.model.SelectedItems)

    @selected_items.setter
    def selected_items(self, value: Iterable[int]) -> None:
        """Sets the selected items"""
        self.model.SelectedItems = tuple(value)  # type: ignore

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        return self.model.Step

    @step.setter
    def step(self, value: int) -> None:
        self.model.Step = value

    @property
    def tab_stop(self) -> bool:
        """Gets/Sets the tab stop property"""
        return self.model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
