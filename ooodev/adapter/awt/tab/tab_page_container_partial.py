from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt.tab import XTabPageContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.awt.tab.uno_control_tab_page_comp import UnoControlTabPageComp

if TYPE_CHECKING:
    from com.sun.star.awt.tab import XTabPageContainerListener
    from com.sun.star.awt.tab import XTabPage
    from ooodev.utils.type_var import UnoInterface


class TabPageContainerPartial:
    """
    Partial class for XTabPageContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTabPageContainer, interface: UnoInterface | None = XTabPageContainer) -> None:
        """
        Constructor

        Args:
            component (XTabPageContainer): UNO Component that implements ``com.sun.star.awt.tab.XTabPageContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTabPageContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTabPageContainer

    def add_tab_page_container_listener(self, listener: XTabPageContainerListener) -> None:
        """
        Adds a listener for the TabPageActiveEvent posted after the tab page was activated.
        """
        self.__component.addTabPageContainerListener(listener)

    def get_tab_page(self, tab_page_index: int) -> UnoControlTabPageComp:
        """
        Returns tab page for the given index.
        """
        result = self.__component.getTabPage(tab_page_index)
        return None if result is None else UnoControlTabPageComp(component=result)  # type: ignore

    def get_tab_page_by_id(self, tab_page_id: int) -> UnoControlTabPageComp:
        """
        Returns tab page for the given ID.
        """
        result = self.__component.getTabPageByID(tab_page_id)
        return None if result is None else UnoControlTabPageComp(component=result)  # type: ignore

    def get_tab_page_count(self) -> int:
        """
        Gets the number of tab pages.
        """
        return self.__component.getTabPageCount()

    def is_tab_page_active(self, tab_page_index: int) -> bool:
        """
        Checks whether a tab page is activated.
        """
        return self.__component.isTabPageActive(tab_page_index)

    def remove_tab_page_container_listener(self, listener: XTabPageContainerListener) -> None:
        """
        Removes a listener previously added with addTabPageListener().
        """
        self.__component.removeTabPageContainerListener(listener)

    @property
    def active_tab_page_id(self) -> int:
        """
        Specifies the ID of the current active tab page.
        """
        return self.__component.ActiveTabPageID

    @active_tab_page_id.setter
    def active_tab_page_id(self, value: int) -> None:
        self.__component.ActiveTabPageID = value

    # endregion XTabPageContainer
