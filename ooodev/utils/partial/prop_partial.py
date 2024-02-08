from __future__ import annotations
from typing import Any, overload
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils import props as mProps
from ooodev.utils import gen_util as gUtil
from ooodev.utils.context.lo_context import LoContext
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.lo_events import observe_events
from ooodev.events.event_singleton import _Events


class PropPartial:
    """
    Property Partial Class.

    If this class is used in a class that inherits from EventsPartial, it will automatically observe events
    for property setting and getting while in the context of this class.
    """

    def __init__(self, component: Any, lo_inst: LoInst):
        """
        Constructor

        Args:
            component (Any): Any Uno component that has properties.
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.
        """
        self.__lo_inst = lo_inst
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
            if isinstance(self, EventsPartial):
                with observe_events(observer=self.event_observer, events=_Events()):
                    result = mProps.Props.get(obj=self.__component, name=name, default=default)
            else:
                result = mProps.Props.get(obj=self.__component, name=name, default=default)
        return result

    def set_property(self, **kwargs: Any) -> None:
        """
        Set property value

        Args:
            **kwargs: Variable length Key value pairs used to set properties.
        """
        with LoContext(self.__lo_inst):
            if isinstance(self, EventsPartial):
                with observe_events(observer=self.event_observer, events=_Events()):
                    mProps.Props.set(self.__component, **kwargs)
            else:
                mProps.Props.set(self.__component, **kwargs)
