from __future__ import annotations
from typing import TYPE_CHECKING, Tuple

import uno
from com.sun.star.frame import XController2
from ooodev.adapter.frame.controller_partial import ControllerPartial
from ooodev.adapter.awt.window_comp import WindowComp

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.ui import XSidebarProvider
    from ooodev.utils.type_var import UnoInterface


class Controller2Partial(ControllerPartial):
    """
    Partial class for XController2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XController2, interface: UnoInterface | None = XController2) -> None:
        """
        Constructor

        Args:
            component (XController2 ): UNO Component that implements ``com.sun.star.frame.XController2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XController2``.
        """

        ControllerPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XController2
    def get_sidebar(self) -> XSidebarProvider:
        """
        Get the sidebar if exists

        **since**

            LibreOffice 5.1
        """
        return self.__component.getSidebar()

    @property
    def creation_arguments(self) -> Tuple[PropertyValue, ...]:
        """
        Denotes the arguments used to create the instance.

        Usually, controllers are created via XModel2.createViewController(), where the caller can pass not only a controller name, but also arguments parameterizing the to-be-created instance. Those arguments used at creation time can subsequently be retrieved using the CreationArguments member.
        """
        ...

    @property
    def component_window(self) -> WindowComp:
        """
        denotes the \"root window\" of the controller.

        If the controller is plugged into a frame, this window acts as the frame's ComponentWindow.
        """
        return WindowComp(self.__component.ComponentWindow)

    @property
    def view_controller_name(self) -> str:
        """
        specifies the view name of the controller.

        A view name is a logical name, which can be used to create views of the same type. The name is meaningful only in conjunction with XModel2.createViewController()
        """
        return self.__component.ViewControllerName

    # endregion XController2
