# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.tab.tab_page_container_events import TabPageContainerEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt.tab import UnoControlTabPageContainer  # service
    from com.sun.star.awt.tab import UnoControlTabPageContainerModel  # service
# endregion imports


class CtlTabPageContainer(DialogControlBase, TabPageContainerEvents):
    """Class for Tab Page Container Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlTabPageContainer) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlTabPageContainer): Tab Page Container Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        TabPageContainerEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_tab_page_container_listener_add_remove
        )

    # endregion init

    # region Lazy Listeners
    def _on_tab_page_container_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTabPageContainerListener(self.events_listener_tab_page_container)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlTabPageContainer:
        return cast("UnoControlTabPageContainer", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlTabPageContainer``"""
        return "com.sun.star.awt.UnoControlTabPageContainer"

    def get_model(self) -> UnoControlTabPageContainerModel:
        """Gets the Model for the control"""
        return cast("UnoControlTabPageContainerModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.TAB_PAGE_CONTAINER``"""
        return DialogControlKind.TAB_PAGE_CONTAINER

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.TAB_PAGE_CONTAINER``"""
        return DialogControlNamedKind.TAB_PAGE_CONTAINER

    # endregion Overrides

    # region Methods
    def is_tab_page_active(self, tab_page_id: int) -> bool:
        """Checks if a tab page is active"""
        return self.view.isTabPageActive(tab_page_id)

    # endregion Methods

    # region Properties
    @property
    def view(self) -> UnoControlTabPageContainer:
        # pylint: disable=no-member
        return cast("UnoControlTabPageContainer", super().view)

    @property
    def model(self) -> UnoControlTabPageContainerModel:
        # pylint: disable=no-member
        return cast("UnoControlTabPageContainerModel", super().model)

    @property
    def tab_page_count(self) -> int:
        """Gets the number of tab pages"""
        return self.view.getTabPageCount()

    @property
    def active_tab_page_id(self) -> int:
        """Gets or sets the active tab page"""
        return self.view.ActiveTabPageID

    @active_tab_page_id.setter
    def active_tab_page_id(self, value: int) -> None:
        self.view.ActiveTabPageID = value

    # endregion Properties
