from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XSearchable
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from .search_descriptor_comp import SearchDescriptorComp

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.util import XSearchDescriptor
    from com.sun.star.uno import XInterface


class SearchablePartial:
    """
    Partial Class XSearchable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSearchable, interface: UnoInterface | None = XSearchable) -> None:
        """
        Constructor

        Args:
            component (XSearchable): UNO Component that implements ``com.sun.star.util.XSearchable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSearchable``.
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

    # region XSearchable
    def create_search_descriptor(self) -> SearchDescriptorComp:
        """
        Creates a Search Descriptor which contains properties that specify a search in this container.

        Returns:
            SearchDescriptorComp: The search descriptor.
        """
        return SearchDescriptorComp(self.__component.createSearchDescriptor())  # type: ignore

    def find_all(self, desc: XSearchDescriptor) -> IndexAccessComp | None:
        """
        Searches the contained texts for all occurrences of whatever is specified.

        Returns:
            IndexAccessComp | None: The found occurrences.
        """
        result = self.__component.findAll(desc)
        return None if result is None else IndexAccessComp(result)

    def find_first(self, desc: XSearchDescriptor) -> XInterface | None:
        """
        Searches the contained texts for the next occurrence of whatever is specified.
        """
        return self.__component.findFirst(desc)

    def find_next(self, start: XInterface, desc: XSearchDescriptor) -> XInterface | None:
        """
        Searches the contained texts for the next occurrence of whatever is specified.
        """
        return self.__component.findNext(start, desc)

    # endregion XSearchable
