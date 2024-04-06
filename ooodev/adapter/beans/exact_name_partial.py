from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XExactName


from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ExactNamePartial:
    """
    Partial class for XExactName.
    """

    def __init__(self, component: XExactName, interface: UnoInterface | None = XExactName) -> None:
        """
        Constructor

        Args:
            component (XExactName): UNO Component that implements ``com.sun.star.container.XExactName`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XExactName``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XExactName

    def get_exact_name(self, approximate_name: str) -> str:
        """
        For example ``getExactName`` could be returned for ``GETEXACTNAME`` when ``GETEXACTNAME`` was used by a case insensitive scripting language.
        """
        return self.__component.getExactName(approximate_name)

    # endregion XExactName


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
        name="ooodev.adapter.beans.exact_name_partial.ExactNamePartial",
        uno_name="com.sun.star.beans.XExactName",
        optional=False,
        init_kind=2,
    )
    return builder
