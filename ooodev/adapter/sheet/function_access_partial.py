from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.sheet import XFunctionAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class FunctionAccessPartial:
    """
    Partial Class for XFunctionAccess.

    .. versionadded:: 0.34.4
    """

    def __init__(self, component: XFunctionAccess, interface: UnoInterface | None = XFunctionAccess) -> None:
        """
        Constructor

        Args:
            component (XFunctionAccess): UNO Component that implements ``com.sun.star.sheet.XFunctionAccess``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFunctionAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XFunctionAccess
    def call_function(self, name: str, *args: Any) -> Any:
        """
        Calls a function and returns the result of the call.

        Each element must be of one of the following types:

        Possible types for the result are:

        - ``int`` or ``float`` for a numeric value.
        - ``str`` for a string value.
        - ``Tuple[Tuple[int,...], ...]`` or ``Tuple[Tuple[float, ...], ...]`` for an array of numeric values.
        - ``Tuple[[str, ...], ...]`` for an array of string values.
        - ``Tuple[[Any, ...], ...]`` for a mixed array, where each element must be of ``None``, ``int``, ``float`` or ``str`` type.


        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``

        Returns:
            Any: the result of the function call.

        Note:
            Possible types for the result are:

            - ``None`` if no result is available.
            - ``float`` for a numeric result.
            - ``str`` for a string result.
            - Tuple[[Any, ...], ...] for an array result, containing ``float`` and ``str`` values.
        """
        return self.__component.callFunction(name, args)

    # endregion XFunctionAccess
