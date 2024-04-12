from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.frame import XLayoutManager
from ooo.dyn.awt.point import Point
from ooo.dyn.awt.size import Size

from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.units.unit_px import UnitPX
from ooodev.utils.data_type.generic_unit_rect import GenericUnitRect
from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
from ooodev.utils.data_type.generic_unit_size import GenericUnitSize
from ooodev.units.unit_convert import UnitLength

if TYPE_CHECKING:
    from com.sun.star.ui import XDockingAreaAcceptor
    from com.sun.star.ui import XUIElement
    from com.sun.star.frame import XFrame
    from ooo.dyn.ui.ui_element_type import UIElementTypeEnum
    from ooo.dyn.ui.docking_area import DockingArea
    from ooodev.utils.type_var import UnoInterface


class LayoutManagerPartial:
    """
    Partial class for XLayoutManager.
    """

    def __init__(self, component: XLayoutManager, interface: UnoInterface | None = XLayoutManager) -> None:
        """
        Constructor

        Args:
            component (XLayoutManager): UNO Component that implements ``com.sun.star.frame.XLayoutManager`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLayoutManager``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLayoutManager
    def attach_frame(self, Frame: XFrame) -> None:
        """
        attaches a com.sun.star.frame.XFrame to a layout manager.

        A layout manager needs a com.sun.star.frame.XFrame to be able to work. Without a it no user interface elements can be created.
        """
        self.__component.attachFrame(Frame)

    def create_element(self, resource_url: str) -> None:
        """
        Creates a new user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be created.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        self.__component.createElement(resource_url)

    def destroy_element(self, resource_url: str) -> None:
        """
        Destroys a user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be created.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        self.__component.destroyElement(resource_url)

    def do_layout(self) -> None:
        """
        Forces a complete new layouting of all user interface elements.
        """
        self.__component.doLayout()

    def dock_all_windows(self, element_type: int | UIElementTypeEnum) -> bool:
        """
        Docks all windows which are member of the provided user interface element type.

        Args:
            element_type (int | UIElementTypeEnum): Specifies the user interface element type.

        Returns:
            bool: Returns ``True`` if all user interface elements of the requested type could be docked, otherwise ``False`` will be returned.

        Hint:
            - ``UIElementTypeEnum`` is an enum and can be imported from ``ooo.dyn.ui.ui_element_type``.
        """
        return self.__component.dockAllWindows(int(element_type))

    def dock_window(self, resource_url: str, docking_area: DockingArea, pos: Point | GenericUnitPoint) -> bool:
        """
        Docks a window based user interface element to a specified docking area.

        Args:
            resource_url (str): Specifies which user interface element should be docked.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
            docking_area (DockingArea): Specifies the docking area where the user interface element should be docked.
            pos (Point): Specifies the position inside the docking area.
                If ``pos`` is a ``GenericUnitPoint`` object, it will be converted pixels first an then to a Point object.

        Returns:
            bool: Returns ``True`` if the user interface element has been docked, otherwise ``False`` will be returned.

        Hint:
            - ``DockingArea`` is an enum and can be imported from ``ooo.dyn.ui.docking_area``.
            - ``Point`` is a struct and can be imported from ``ooo.dyn.awt.point``.
        """
        if isinstance(pos, GenericUnitPoint):
            px_unit = pos.convert_to(UnitLength.PX)
            p = px_unit.get_uno_point()
        else:
            p = pos
        return self.__component.dockWindow(resource_url, docking_area, p)  # type: ignore

    def float_window(self, resource_url: str) -> bool:
        """
        Forces a window based user interface element to float.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.

        Returns:
            bool: Returns ``True`` if the user interface element has been docked, otherwise ``False`` will be returned.
        """
        return self.__component.floatWindow(resource_url)

    def get_current_docking_area(self) -> GenericUnitRect[UnitPX, float]:
        """
        Provides the current docking area size of the layout manager.

        Returns:
            GenericUnitRect[UnitPX, float]: The current docking area size represented as pixels
        """
        rect = self.__component.getCurrentDockingArea()
        return GenericUnitRect(
            UnitPX(float(rect.X)), UnitPX(float(rect.Y)), UnitPX(float(rect.Width)), UnitPX(float(rect.Height))
        )

    def get_docking_area_acceptor(self) -> XDockingAreaAcceptor:
        """
        retrieves the current docking area acceptor that controls the border space of the frame's container window.

        A docking area acceptor retrieved by this method is owned by the layout manager. It is not allowed to dispose this object, it will be destroyed on reference count!
        """
        return self.__component.getDockingAreaAcceptor()

    def get_element(self, resource_url: str) -> XUIElement:
        """
        retrieves a user interface element which has been created before.

        The layout manager instance is owner of the returned user interface element.
        That means that the life time of the user interface element is controlled by the layout manager.
        It can be disposed at every time!

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.getElement(resource_url)

    def get_element_pos(self, resource_url: str) -> GenericUnitPoint[UnitPX, float]:
        """
        Retrieves the current pixel position of a window based user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.

        Returns:
            GenericUnitPoint[UnitPX, float]: The current position of the user interface element represented as pixels.
        """
        p = self.__component.getElementPos(resource_url)
        return GenericUnitPoint(UnitPX(float(p.X)), UnitPX(float(p.Y)))

    def get_element_size(self, resource_url: str) -> GenericUnitSize[UnitPX, float]:
        """
        Retrieves the current size of a window based user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        sz = self.__component.getElementSize(resource_url)
        return GenericUnitSize(UnitPX(float(sz.Width)), UnitPX(float(sz.Height)))

    def get_elements(self) -> Tuple[XUIElement, ...]:
        """
        Retrieves all user interface elements which are currently instantiated.

        The layout manager instance is owner of the returned user interface elements.
        That means that the life time of the user interface elements is controlled by the layout manager.
        They can be disposed at every time!
        """
        return self.__component.getElements()

    def hide_element(self, resource_url: str) -> bool:
        """
        Hides a user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.hideElement(resource_url)

    def is_element_docked(self, resource_url: str) -> bool:
        """
        Retrieves the current docking state of a window based user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.isElementDocked(resource_url)

    def is_element_floating(self, resource_url: str) -> bool:
        """
        Retrieves the current floating state of a window based user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.isElementFloating(resource_url)

    def is_element_locked(self, resource_url: str) -> bool:
        """
        Retrieves the current lock state of a window based user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.isElementLocked(resource_url)

    def is_element_visible(self, resource_url: str) -> bool:
        """
        Retrieves the current visibility state of a window based user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.isElementVisible(resource_url)

    def is_visible(self) -> bool:
        """
        Retrieves the visibility state of a layout manager.

        A layout manager can be set to invisible state to force it to hide all of its user interface elements.
        If another component wants to use the window for its own user interface elements it can use this function.
        This function is normally used to implement inplace editing.
        """
        return self.__component.isVisible()

    def lock(self) -> None:
        """
        Prohibit all layout updates until unlock is called again.

        This call can be used to speed up the creation process of several user interface elements.
        Otherwise the layout manager would calculate the layout for every creation.
        """
        self.__component.lock()

    def lock_window(self, resource_url: str) -> bool:
        """
        Locks a window based user interface element if it's in a docked state.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.lockWindow(resource_url)

    def request_element(self, resource_url: str) -> bool:
        """
        Request to make a user interface element visible if it is not in hidden state.

        If a user interface element should forced to the visible state ``XLayoutManager.showElement()`` should be used.
        This function can be used for context dependent elements which should respect the current visibility state.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
        """
        return self.__component.requestElement(resource_url)

    def reset(self) -> None:
        """
        Resets the layout manager and remove all of its internal user interface elements.

        This call should be handled with care as all user interface elements will be destroyed and the layout manager is reset
        to a state after a ``attach_frame()`` has been made. That means an attached frame which has been set by ``attach_frame()`` is not released.
        The layout manager itself calls reset after a component has been attached or reattached to a frame.
        """
        self.__component.reset()

    def set_docking_area_acceptor(self, docking_area_acceptor: XDockingAreaAcceptor) -> None:
        """
        sets a docking area acceptor that controls the border space of the frame's container window.

        A docking area acceptor decides if the layout manager can use requested border space for docking windows.
        If the acceptor denies the requested space the layout manager automatically set all docked windows into
        floating state and will not use this space for docking.After setting a docking area acceptor the object is
        owned by the layout manager. It is not allowed to dispose this object, it will be destroyed on reference count!
        """
        self.__component.setDockingAreaAcceptor(docking_area_acceptor)

    def set_element_pos(self, resource_url: str, pos: Point | GenericUnitPoint) -> None:
        """
        Sets a new position for a window based user interface element.

        It is up to the layout manager to decide if the user interface element can be moved.
        The new position can be retrieved by calling ``get_element_pos()``.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
            pos (Point | GenericUnitPoint): Specifies the new position of the user interface element.
                If ``pos`` is a ``GenericUnitPoint`` object, it will be converted pixels first an then to a Point object.
        """
        if isinstance(pos, GenericUnitPoint):
            px_unit = pos.convert_to(UnitLength.PX)
            p = px_unit.get_uno_point()
        else:
            p = pos
        self.__component.setElementPos(resource_url, p)

    def set_element_pos_size(
        self, resource_url: str, pos: Point | GenericUnitPoint, size: Size | GenericUnitSize
    ) -> None:
        """
        sets a new position and size for a window based user interface element.

        It is up to the layout manager to decide if the user interface element can be moved and resized.
        The new position and size can be retrieved by calling ``get_element_pos()`` and ``get_element_size()``.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.
            pos (Point | GenericUnitPoint): Specifies the new position of the user interface element.
                If ``pos`` is a ``GenericUnitPoint`` object, it will be converted pixels first an then to a Point object.
            size (Size | GenericUnitSize): Specifies the new size of the user interface element.
                If ``pos`` is a ``GenericUnitSize`` object, it will be converted pixels first an then to a Point object.
        """
        if isinstance(pos, GenericUnitPoint):
            px_unit = pos.convert_to(UnitLength.PX)
            p = px_unit.get_uno_point()
        else:
            p = pos

        if isinstance(size, GenericUnitSize):
            px_unit = size.convert_to(UnitLength.PX)
            sz = px_unit.get_uno_size()
        else:
            sz = size
        self.__component.setElementPosSize(resource_url, p, sz)

    def set_element_size(self, resource_url: str, size: Size | GenericUnitSize) -> None:
        """
        Sets a new size for a window based user interface element.

        It is up to the layout manager to decide if the user interface element can be resized.
        The new size can be retrieved by calling ``get_element_size()``.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.

            size (Size | GenericUnitSize): Specifies the new size of the user interface element.
                If ``pos`` is a ``GenericUnitSize`` object, it will be converted pixels first an then to a Point object.
        """
        if isinstance(size, GenericUnitSize):
            px_unit = size.convert_to(UnitLength.PX)
            sz = px_unit.get_uno_size()
        else:
            sz = size
        self.__component.setElementSize(resource_url, sz)

    def set_visible(self, visible: bool) -> None:
        """
        Sets the layout manager to invisible state and hides all user interface elements.

        A layout manager can be set to invisible state to force it to hide all of its user interface elements.
        If another component wants to use the window for its own user interface elements it can use this function.
        This function is normally used to implement inplace editing.
        """
        self.__component.setVisible(visible)

    def show_element(self, resource_url: str) -> bool:
        """
        Shows a user interface element.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.

        Returns:
            bool: Returns ``True`` if the user interface element has been shown, otherwise ``False`` will be returned.
        """
        return self.__component.showElement(resource_url)

    def unlock(self) -> None:
        """
        Permit layout updates again.

        This function should be called to permit layout updates. The layout manager starts to calculate the new layout after this call.
        """
        self.__component.unlock()

    def unlock_window(self, resource_url: str) -> bool:
        """
        Unlocks a window based user interface element if it's in a docked state.

        Args:
            resource_url (str): Specifies which user interface element should be floated.
                A resource URL must meet the following syntax: ``private:resource/$type/$name``.
                It is only allowed to use ASCII characters for type and name.

        Returns:
            bool: Returns ``True`` if the user interface element has been unlocked, otherwise ``False`` will be returned.
        """
        return self.__component.unlockWindow(resource_url)

    # endregion XLayoutManager


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.frame.XLayoutManager", False)
    return builder
