# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.container.container_events import ContainerEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import info as mInfo

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt.tab import UnoControlTabPage  # service
    from com.sun.star.awt.tab import UnoControlTabPageModel  # service
# endregion imports


class CtlTabPage(CtlBase, ContainerEvents):
    """Class for Tab Page Control"""

    # region init
    def __init__(self, ctl: UnoControlTabPage) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlTabPage): Tab Page Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ContainerEvents.__init__(self, trigger_args=generic_args, cb=self._on_container_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_container_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addContainerListener(self.events_listener_container)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlTabPage:
        return cast("UnoControlTabPage", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlTabPage``"""
        return "com.sun.star.awt.UnoControlTabPage"

    def get_model(self) -> UnoControlTabPageModel:
        """Gets the Model for the control"""
        return cast("UnoControlTabPageModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Methods

    # endregion Methods

    # region Properties
    @property
    def view(self) -> UnoControlTabPage:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlTabPageModel:
        return self.get_model()

    @property
    def horizontal_scrollbar(self) -> bool:
        """
        Gets or sets if a horizontal scrollbar should be added to the dialog.

        Returns ``False`` if LibreOffice version is less than ``7.1``

        **since**

            LibreOffice ``7.1``
        """
        if mInfo.Info.version_info < (7, 1):
            return False
        return self.model.HScroll

    @horizontal_scrollbar.setter
    def horizontal_scrollbar(self, value: bool) -> None:
        """Sets the horizontal scrollbar"""
        if mInfo.Info.version_info < (7, 1):
            return
        self.model.HScroll = value

    @property
    def vertical_scrollbar(self) -> bool:
        """
        Gets or sets if a vertical scrollbar should be added to the dialog.

        Returns ``False`` if LibreOffice version is less than ``7.1``

        **since**

            LibreOffice ``7.1``
        """
        if mInfo.Info.version_info < (7, 1):
            return False
        return self.model.VScroll

    @vertical_scrollbar.setter
    def vertical_scrollbar(self, value: bool) -> None:
        if mInfo.Info.version_info < (7, 1):
            return
        self.model.VScroll = value

    @property
    def scroll_height(self) -> int:
        """
        Gets or sets the total height of the scrollable dialog content

        Returns ``-1`` if LibreOffice version is less than ``7.1``

        **since**

            LibreOffice ``7.1``
        """
        if mInfo.Info.version_info < (7, 1):
            return -1
        return self.model.ScrollHeight

    @scroll_height.setter
    def scroll_height(self, value: int) -> None:
        """Sets the scroll height"""
        if mInfo.Info.version_info < (7, 1):
            return
        self.model.ScrollHeight = value

    @property
    def scroll_left(self) -> int:
        """
        Gets or sets the horizontal position of the scrolled dialog content

        Returns ``-1`` if LibreOffice version is less than ``7.1``

        **since**

            LibreOffice ``7.1``
        """
        if mInfo.Info.version_info < (7, 1):
            return -1
        return self.model.ScrollLeft

    @scroll_left.setter
    def scroll_left(self, value: int) -> None:
        """Sets the scroll left"""
        if mInfo.Info.version_info < (7, 1):
            return
        self.model.ScrollLeft = value

    @property
    def scroll_top(self) -> int:
        """
        Gets or sets the vertical position of the scrolled dialog content

        Returns ``-1`` if LibreOffice version is less than ``7.1``

        **since**

            LibreOffice ``7.1``
        """
        if mInfo.Info.version_info < (7, 1):
            return -1
        return self.model.ScrollTop

    @scroll_top.setter
    def scroll_top(self, value: int) -> None:
        """Sets the scroll top"""
        if mInfo.Info.version_info < (7, 1):
            return
        self.model.ScrollTop = value

    @property
    def scroll_width(self) -> int:
        """
        Gets or sets the total width of the scrollable dialog content

        Returns ``-1`` if LibreOffice version is less than ``7.1``

        **since**

            LibreOffice ``7.1``
        """
        if mInfo.Info.version_info < (7, 1):
            return -1
        return self.model.ScrollWidth

    @scroll_width.setter
    def scroll_width(self, value: int) -> None:
        """Sets the scroll width"""
        if mInfo.Info.version_info < (7, 1):
            return
        self.model.ScrollWidth = value

    @property
    def title(self) -> str:
        """Gets or sets the title of the tab page"""
        return self.model.Title

    @title.setter
    def title(self, value: str) -> None:
        """Sets the title of the tab page"""
        self.model.Title = value

    # endregion Properties
