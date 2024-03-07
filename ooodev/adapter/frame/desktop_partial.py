from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.frame import XDesktop

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.frame import XTerminateListener
    from com.sun.star.container import XEnumerationAccess
    from com.sun.star.lang import XComponent
    from com.sun.star.frame import XFrame
    from ooodev.utils.type_var import UnoInterface


class DesktopPartial:
    """
    Partial class for XDesktop.
    """

    def __init__(self, component: XDesktop, interface: UnoInterface | None = XDesktop) -> None:
        """
        Constructor

        Args:
            component (XDesktop): UNO Component that implements ``com.sun.star.frame.XDesktop`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDesktop``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDesktop
    def add_terminate_listener(self, listener: XTerminateListener) -> None:
        """
        Registers an event listener to the desktop, which is called when the desktop is queried to terminate, and when it really terminates.
        """
        self.__component.addTerminateListener(listener)

    def get_components(self) -> XEnumerationAccess:
        """
        Provides read access to collection of all currently loaded components inside the frame tree

        The component is, by definition, the model of the control which is loaded into a frame, or if no model exists, into the control itself. The service Components which is available from this method is a collection of all components of the desktop which are open within a frame of the desktop.
        """
        return self.__component.getComponents()

    def get_current_component(self) -> XComponent:
        """
        Provides read access to the component inside the tree which has the UI focus

        Normally, the component is the model part of the active component. If no model exists it is the active controller (view) itself.
        """
        return self.__component.getCurrentComponent()

    def get_current_frame(self) -> XFrame:
        """
        provides read access to the frame which contains the current component
        """
        return self.__component.getCurrentFrame()

    def remove_terminate_listener(self, listener: XTerminateListener) -> None:
        """
        Un-registers an event listener for termination events.
        """
        self.__component.removeTerminateListener(listener)

    def terminate(self) -> bool:
        """
        Tries to terminate the desktop.

        First, every terminate listener is called by this ``XTerminateListener.queryTermination()`` method.
        Throwing of a ``TerminationVetoException`` can break the termination process and the listener how has done that will be
        the new ``controller`` of the desktop lifetime. Should try to terminate it by itself after his own processes will be finished.
        If nobody disagree with the termination request, every listener will be called by his ``XTerminateListener.notifyTermination()`` method.
        """
        return self.__component.terminate()

    # endregion XDesktop
