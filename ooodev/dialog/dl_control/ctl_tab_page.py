# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.container.container_events import ContainerEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import info as mInfo
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt.tab import UnoControlTabPage  # service
    from com.sun.star.awt.tab import UnoControlTabPageModel  # service
# endregion imports


class CtlTabPage(DialogControlBase, ContainerEvents):
    """Class for Tab Page Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlTabPage) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlTabPage): Tab Page Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ContainerEvents.__init__(self, trigger_args=generic_args, cb=self._on_container_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_container_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addContainerListener(self.events_listener_container)
        event.remove_callback = True

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

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.TAB_PAGE``"""
        return DialogControlKind.TAB_PAGE

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.TAB_PAGE``"""
        return DialogControlNamedKind.TAB_PAGE

    # endregion Overrides

    # region Methods

    # endregion Methods

    # region Properties
    @property
    def model(self) -> UnoControlTabPageModel:
        # pylint: disable=no-member
        return cast("UnoControlTabPageModel", super().model)

    @property
    def view(self) -> UnoControlTabPage:
        # pylint: disable=no-member
        return cast("UnoControlTabPage", super().view)

    @property
    def horizontal_scrollbar(self) -> bool:
        """
        Gets or sets if a horizontal scrollbar should be added to the dialog.

        Returns ``False`` if LibreOffice version is less than ``7.1``

        **since**

            LibreOffice ``7.1``
        """
        return False if mInfo.Info.version_info < (7, 1) else self.model.HScroll

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
        return False if mInfo.Info.version_info < (7, 1) else self.model.VScroll

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
        return -1 if mInfo.Info.version_info < (7, 1) else self.model.ScrollHeight

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
        return -1 if mInfo.Info.version_info < (7, 1) else self.model.ScrollLeft

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
        return -1 if mInfo.Info.version_info < (7, 1) else self.model.ScrollTop

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
        return -1 if mInfo.Info.version_info < (7, 1) else self.model.ScrollWidth

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
