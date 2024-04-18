from __future__ import annotations
from typing import Any, Generic, TypeVar, TYPE_CHECKING
import uno

from com.sun.star.container import XHierarchicalNameAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


T = TypeVar("T")


class HierarchicalNameAccessPartial(Generic[T]):
    """
    Partial class for XHierarchicalNameAccess.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XHierarchicalNameAccess, interface: UnoInterface | None = XHierarchicalNameAccess
    ) -> None:
        """
        Constructor

        Args:
            component (XHierarchicalNameAccess): UNO Component that implements ``com.sun.star.container.XHierarchicalNameAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XHierarchicalNameAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XHierarchicalNameAccess
    def get_by_hierarchical_name(self, name: str) -> T:
        """

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.getByHierarchicalName(name)

    def has_by_hierarchical_name(self, name: str) -> bool:
        """
        In many cases, the next call is XNameAccess.getByName(). You should optimize this case.
        """
        return self.__component.hasByHierarchicalName(name)

    # endregion XHierarchicalNameAccess


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.add_import(
        name="ooodev.adapter.container.hierarchical_name_access_partial.HierarchicalNameAccessPartial",
        uno_name="com.sun.star.container.XHierarchicalNameAccess",
        optional=False,
        init_kind=2,
    )
    return builder
