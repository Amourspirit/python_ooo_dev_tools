from __future__ import annotations
from typing import Any, overload
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils import props as mProps
from ooodev.utils import gen_util as gUtil
from ooodev.utils.context.lo_context import LoContext


class PropPartial:
    def __init__(self, component: Any, lo_inst: LoInst):
        self.__lo_inst = lo_inst  # may be used in future
        self.__component = component

    @overload
    def get_property(self, name: str) -> Any:
        """
        Get property value

        Args:
            name (str): Property Name.

        Returns:
            Any: Property value or default.
        """
        ...

    @overload
    def get_property(self, name: str, default: Any) -> Any:
        """
        Get property value

        Args:
            name (str): Property Name.
            default (Any, optional): Return value if property value is ``None``.

        Returns:
            Any: Property value or default.
        """
        ...

    def get_property(self, name: str, default: Any = gUtil.NULL_OBJ) -> Any:
        """
        Get property value

        Args:
            name (str): Property Name.
            default (Any, optional): Return value if property value is ``None``.

        Returns:
            Any: Property value or default.
        """
        with LoContext(self.__lo_inst):
            result = mProps.Props.get(obj=self.__component, name=name, default=default)
        return result

    def set_property(self, **kwargs: Any) -> None:
        """
        Set property value

        Args:
            **kwargs: Variable length Key value pairs used to set properties.
        """
        with LoContext(self.__lo_inst):
            mProps.Props.set(self.__component, **kwargs)
