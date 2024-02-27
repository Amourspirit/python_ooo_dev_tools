from __future__ import annotations
from typing import Any, cast, Iterable, TYPE_CHECKING
import uno

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.utils.kind.border_kind import BorderKind as BorderKind


from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import ComboBox as ControlModel  # service
    from com.sun.star.form.control import ComboBox as ControlView  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlComboBox(FormCtlBase, ActionEvents, ItemEvents, TextEvents, ResetEvents):
    """``com.sun.star.form.component.ComboBox`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.ComboBox`` service.
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
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        ResetEvents.__init__(self, trigger_args=generic_args, cb=self._on_reset_add_remove)

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addActionListener(self.events_listener_action)
        event.remove_callback = True

    def _on_item_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addItemListener(self.events_listener_item)
        event.remove_callback = True

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    def _on_reset_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addResetListener(self.events_listener_reset)
        event.remove_callback = True

    # endregion Lazy Listeners

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

    # endregion Methods

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
        return FormComponentKind.COMBO_BOX

    # endregion Overrides

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
        """Gets/Sets the combobox has a drop down button."""
        return self.model.Dropdown

    @drop_down.setter
    def drop_down(self, value: bool) -> None:
        self.model.Dropdown = value

    @property
    def max_text_len(self) -> int:
        """Gets/Sets the maximum character count."""
        return self.model.MaxTextLen

    @max_text_len.setter
    def max_text_len(self, value: int) -> None:
        self.model.MaxTextLen = value

    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def read_only(self) -> bool:
        """Gets/Sets the read-only property"""
        return self.model.ReadOnly

    @read_only.setter
    def read_only(self, value: bool) -> None:
        """Sets the read-only property"""
        self.model.ReadOnly = value

    @property
    def tab_stop(self) -> bool:
        """Gets/Sets the tab stop property"""
        return self.model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
