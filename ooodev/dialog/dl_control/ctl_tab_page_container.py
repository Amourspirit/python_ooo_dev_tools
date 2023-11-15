# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.tab.tab_page_container_events import TabPageContainerEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

from .ctl_base import DialogControlBase

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
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTabPageContainerListener(self.events_listener_tab_page_container)
        self._add_listener(key)

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

    # endregion Overrides

    # region Methods
    def is_tab_page_active(self, tab_page_id: int) -> bool:
        """Checks if a tab page is active"""
        return self.view.isTabPageActive(tab_page_id)

    # endregion Methods

    # region Properties
    @property
    def view(self) -> UnoControlTabPageContainer:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlTabPageContainerModel:
        return self.get_model()

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
