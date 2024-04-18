from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.task import XStatusIndicatorSupplier

from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.task import XStatusIndicator
    from ooodev.utils.type_var import UnoInterface


class StatusIndicatorSupplierPartial:
    """
    Partial class for XStatusIndicatorSupplier.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XStatusIndicatorSupplier, interface: UnoInterface | None = XStatusIndicatorSupplier
    ) -> None:
        """
        Constructor

        Args:
            component (XStatusIndicatorSupplier): UNO Component that implements ``com.sun.star.container.XStatusIndicatorSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStatusIndicatorSupplier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCellRangesAccess
    def get_status_indicator(self) -> XStatusIndicator:
        """
        use XStatusIndicatorFactory.createStatusIndicator() instead of this
        """
        return self.__component.getStatusIndicator()

    # endregion XCellRangesAccess


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
    builder.auto_add_interface("com.sun.star.task.XStatusIndicatorSupplier", False)
    return builder
