from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XFilter

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import PropertyValue


class FilterPartial:
    """
    Partial class for XFilter.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFilter, interface: UnoInterface | None = XFilter) -> None:
        """
        Constructor

        Args:
            component (XFilter): UNO Component that implements ``com.sun.star.document.XFilter`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFilter``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XFilter
    def cancel(self) -> None:
        """Cancels the export process."""
        self.__component.cancel()

    def filter(self, *args: PropertyValue) -> None:
        """
        Filter the document.

        Args:
            url (str): The URL of the graphic file to be exported.
            args (Tuple[PropertyValue, ...]): The arguments to be used for the export filter.
        """
        # desc = uno.Any("[]com.sun.star.beans.PropertyValue", args)  # type: ignore
        # uno.invoke(self.__component, "filter", (desc,))
        self.__component.filter(args)

    # endregion XFilter
