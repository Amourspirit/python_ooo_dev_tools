from __future__ import annotations
from typing import Any, overload
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils import props as mProps
from ooodev.utils import gen_util as gUtil


class PropPartial:
    def __init__(self, component: Any, lo_inst: LoInst):
        self.__lo_inst = lo_inst  # may be used in future
        self.__component = component

    @overload
    def get_property(self, name: str) -> Any:
        ...

    @overload
    def get_property(self, name: str, default: Any) -> Any:
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
        return mProps.Props.get(obj=self.__component, name=name, default=default)

    def set_property(self, **kwargs: Any) -> None:
        """
        Set property value

        Args:
            **kwargs: Variable length Key value pairs used to set properties.
        """
        mProps.Props.set(self.__component, **kwargs)
