# region imports
from __future__ import annotations
from typing import Any, cast, Iterable, TYPE_CHECKING, Tuple
import contextlib
import uno

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_list_box_model_partial import UnoControlListBoxModelPartial
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlListBox  # service
    from com.sun.star.awt import UnoControlListBoxModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_list_box import ModelListBox
    from ooodev.dialog.dl_control.view.view_list_box import ViewListBox
# endregion imports


class CtlListBox(DialogControlBase, UnoControlListBoxModelPartial, ActionEvents, ItemEvents):
    """Class for ListBox Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlListBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlListBox): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlListBoxModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlListBox({self.name})"
        return "CtlListBox"

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addActionListener(self.events_listener_action)
        event.remove_callback = True

    def _on_item_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addItemListener(self.events_listener_item)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlListBox:
        return cast("UnoControlListBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlListBox``"""
        return "com.sun.star.awt.UnoControlListBox"

    def get_model(self) -> UnoControlListBoxModel:
        """Gets the Model for the control"""
        return cast("UnoControlListBoxModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.LIST_BOX``"""
        return DialogControlKind.LIST_BOX

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.LIST_BOX``"""
        return DialogControlNamedKind.LIST_BOX

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
        ctl_props = self.get_control_props()
        if ctl_props is None:
            return
        uno_strings = uno.Any("[]string", tuple(data))  # type: ignore
        uno.invoke(ctl_props, "setPropertyValue", ("StringItemList", uno_strings))  # type: ignore

    # endregion Methods

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlListBox":
        """
        Creates a new instance of the control.

        Keyword arguments are optional.
        Extra Keyword args are passed to the control as property values.

        Args:
            win (XWindowPeer): Parent Window

        Keyword Args:
            x (int, UnitT, optional): X Position in Pixels or UnitT.
            y (int, UnitT, optional): Y Position in Pixels or UnitT.
            width (int, UnitT, optional): Width in Pixels or UnitT.
            height (int, UnitT, optional): Height in Pixels or UnitT.

        Returns:
            CtlListBox: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlListBoxModel", win, **kwargs)
        return CtlListBox(ctl=ctrl)

    # endregion Static Methods

    # region Properties

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
    def model(self) -> UnoControlListBoxModel:
        # pylint: disable=no-member
        return cast("UnoControlListBoxModel", super().model)

    @property
    def model_ex(self) -> ModelListBox:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_list_box import ModelListBox

            self._model_ex = ModelListBox(self.model)
        return self._model_ex

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
    def view(self) -> UnoControlListBox:
        # pylint: disable=no-member
        return cast("UnoControlListBox", super().view)

    @property
    def view_ex(self) -> ViewListBox:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_list_box import ViewListBox

            self._view_ex = ViewListBox(self.view)
        return self._view_ex

    # item_count was renamed to list_count in 0.13.2
    item_count = list_count
    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_list_box import ModelListBox
    from ooodev.dialog.dl_control.view.view_list_box import ViewListBox
