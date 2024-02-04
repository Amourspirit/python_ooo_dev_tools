from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.view import XSelectionSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.view import XSelectionChangeListener
    from ooodev.utils.type_var import UnoInterface


class SelectionSupplierPartial:
    """
    Partial class for XSelectionSupplier.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSelectionSupplier, interface: UnoInterface | None = XSelectionSupplier) -> None:
        """
        Constructor

        Args:
            component (XSelectionSupplier ): UNO Component that implements ``com.sun.star.view.XSelectionSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSelectionSupplier``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XSelectionSupplier
    def add_selection_change_listener(self, listener: XSelectionChangeListener) -> None:
        """
        Adds a selection change listener.

        Args:
            listener (XSelectionChangeListener): The listener to be added.
        """
        self.__component.addSelectionChangeListener(listener)

    def get_selection(self) -> Any:
        """
        Returns the current selection.

        The selection is either specified by an object which is contained in the component to which the view belongs,
        or it is an interface of a collection which contains such objects.
        """
        return self.__component.getSelection()

    def remove_selection_change_listener(self, listener: XSelectionChangeListener) -> None:
        """
        Removes a selection change listener.

        Args:
            listener (XSelectionChangeListener): The listener to be removed.
        """
        self.__component.removeSelectionChangeListener(listener)

    def select(self, selection: Any) -> bool:
        """
        selects the object represented by ``selection`` if it is known and selectable in this object.

        Args:
            selection (Any): The selection to be set.

        Returns:
            bool: ``True`` if the selection could be set, otherwise ``False``.
        """
        return self.__component.select(selection)

    # endregion XSelectionSupplier
