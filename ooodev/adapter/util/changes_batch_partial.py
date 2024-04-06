from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.util import XChangesBatch

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.util import ElementChange
    from ooodev.utils.type_var import UnoInterface


class ChangesBatchPartial:
    """
    Partial Class XChangesBatch.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChangesBatch, interface: UnoInterface | None = XChangesBatch) -> None:
        """
        Constructor

        Args:
            component (XChangesBatch): UNO Component that implements ``com.sun.star.util.XChangesBatch`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChangesBatch``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XChangesBatch
    def commit_changes(self) -> None:
        """
        commits any pending changes.

        The exact action depends on the concrete service.

        Raises:
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        return self.__component.commitChanges()

    def get_pending_changes(self) -> Tuple[ElementChange, ...]:
        """
        queries for any pending changes that can be committed.
        """
        return self.__component.getPendingChanges()

    def has_pending_changes(self) -> bool:
        """
        checks whether this object has any pending changes that can be committed.
        """
        return self.__component.hasPendingChanges()

    # endregion XCloneable


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
    builder.auto_add_interface("com.sun.star.util.XChangesBatch", False)
    return builder
