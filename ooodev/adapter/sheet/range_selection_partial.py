from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XRangeSelection

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.sheet import XRangeSelectionListener
    from com.sun.star.sheet import XRangeSelectionChangeListener
    from ooodev.utils.type_var import UnoInterface


class RangeSelectionPartial:
    """
    Partial Class for XRangeSelection.

    .. versionadded:: 0.20.0
    """

    def __init__(self, component: XRangeSelection, interface: UnoInterface | None = XRangeSelection) -> None:
        """
        Constructor

        Args:
            component (XRangeSelection): UNO Component that implements ``com.sun.star.sheet.XRangeSelection``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XRangeSelection``.
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

    # region XRangeSelection
    def abort_range_selection(self) -> None:
        """
        Aborts the current range selection.
        """
        self.__component.abortRangeSelection()

    def add_range_selection_change_listener(self, listener: XRangeSelectionChangeListener) -> None:
        """
        Adds the specified listener to receive events when the selection changes.

        Args:
            listener (XRangeSelectionChangeListener): The listener to add.
        """
        self.__component.addRangeSelectionChangeListener(listener)

    def add_range_selection_listener(self, listener: XRangeSelectionListener) -> None:
        """
        Adds the specified listener to receive events when the selection changes.

        Args:
            listener (XRangeSelectionListener): The listener to add.
        """
        self.__component.addRangeSelectionListener(listener)

    def remove_range_selection_change_listener(self, listener: XRangeSelectionChangeListener) -> None:
        """
        Removes the specified listener so it does not receive events when the selection changes.

        Args:
            listener (XRangeSelectionChangeListener): The listener to remove.
        """
        self.__component.removeRangeSelectionChangeListener(listener)

    def remove_range_selection_listener(self, listener: XRangeSelectionListener) -> None:
        """
        Removes the specified listener so it does not receive events when the selection changes.

        Args:
            listener (XRangeSelectionListener): The listener to remove.
        """
        self.__component.removeRangeSelectionListener(listener)

    def start_range_selection(self, args: tuple[PropertyValue, ...]) -> None:
        """
        Starts a range selection.

        Args:
            args (tuple[PropertyValue, ...]): Specifies how the range selection is done.
        """
        self.__component.startRangeSelection(args)

    # endregion XRangeSelection
