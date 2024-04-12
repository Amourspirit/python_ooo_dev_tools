from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

import uno
from ooo.dyn.frame.command_group import CommandGroupEnum
from com.sun.star.frame import XDispatchInformationProvider

from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.frame import DispatchInformation
    from ooodev.utils.type_var import UnoInterface


class DispatchInformationProviderPartial:
    """
    Partial class for XDispatchInformationProvider.
    """

    def __init__(
        self, component: XDispatchInformationProvider, interface: UnoInterface | None = XDispatchInformationProvider
    ) -> None:
        """
        Constructor

        Args:
            component (XDispatchInformationProvider): UNO Component that implements ``com.sun.star.frame.XDispatchInformationProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDispatchInformationProvider``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDispatchInformationProvider
    def get_configurable_dispatch_information(self, command_group: int) -> Tuple[DispatchInformation, ...]:
        """
        returns additional information about supported commands of a given command group.
        """
        return tuple(self.__component.getConfigurableDispatchInformation(command_group))

    def get_supported_command_groups(self) -> Tuple[CommandGroupEnum, ...]:
        """
        Returns all supported command groups.

        Returns:
            Tuple[CommandGroupEnum, ...]: Supported command groups.

        Note:
            If you want to get the groups as a tuple of integers call ``component.getSupportedCommandGroups()``.

        Hint:
            - ``CommandGroupEnum`` can be imported from ``ooo.dyn.frame.command_group``
        """
        groups = self.__component.getSupportedCommandGroups()  # type: ignore
        return tuple(CommandGroupEnum(group) for group in groups)

    # endregion XDispatchInformationProvider


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.frame.XDispatchProvider", False)
    return builder
