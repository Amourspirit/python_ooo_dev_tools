from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XToolkit

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import Rectangle  # struct
    from com.sun.star.awt import WindowDescriptor  # struct
    from com.sun.star.awt import XDevice
    from com.sun.star.awt import XRegion
    from com.sun.star.awt import XWindowPeer
    from ooodev.utils.type_var import UnoInterface


class ToolkitPartial:
    """
    Partial class for XToolkit.
    """

    def __init__(self, component: XToolkit, interface: UnoInterface | None = XToolkit) -> None:
        """
        Constructor

        Args:
            component (XToolkit): UNO Component that implements ``com.sun.star.awt.XToolkit`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XToolkit``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XToolkit
    def create_region(self) -> XRegion:
        """
        Creates a region.
        """
        return self.__component.createRegion()

    def create_screen_compatible_device(self, width: int, height: int) -> XDevice:
        """
        Creates a virtual device that is compatible with the screen.
        """
        return self.__component.createScreenCompatibleDevice(width, height)

    def create_window(self, descriptor: WindowDescriptor) -> XWindowPeer:
        """
        Creates a new window using the given descriptor.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createWindow(descriptor)

    def create_windows(self, *descriptors: WindowDescriptor) -> Tuple[XWindowPeer, ...]:
        """
        Gets a sequence of windows which are newly created using the given descriptors.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createWindows(descriptors)

    def get_desktop_window(self) -> XWindowPeer:
        """
        Gets the desktop window.
        """
        return self.__component.getDesktopWindow()

    def get_work_area(self) -> Rectangle:
        """
        For LibreOffice versions < 4.1, this method just returned an empty rectangle.

        After that, it started returning a valid value.
        """
        return self.__component.getWorkArea()

    # endregion XToolkit
