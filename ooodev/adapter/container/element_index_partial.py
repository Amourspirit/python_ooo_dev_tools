from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.container import XNamed
    from ooodev.adapter.container.name_index_t import NameIndexT
else:
    NameIndexT = object


class ElementIndexPartial:
    """Adds methods for getting the element index by name."""

    def __init__(self, component: NameIndexT):
        self.__component = component

    def get_index_by_name(self, name: str) -> int:
        """
        Gets the element index by name.

        Args:
            name (str): The name of the element.

        Returns:
            int: The index of the element if found; Otherwise ``-1``.
        """

        index = -1
        for i in range(self.__component.get_count()):
            itm = cast("XNamed", self.__component.get_by_index(i))
            if hasattr(itm, "component"):
                with contextlib.suppress(AttributeError):
                    # get name from component because some objects will return a Python class not a UNO object.
                    # For instance CalcForms will return a CalcForm from get_by_index()
                    if cast(Any, itm).component.getName() == name:
                        index = i
                        break
            else:
                with contextlib.suppress(AttributeError):
                    if itm.getName() == name:
                        index = i
                        break
        return index
