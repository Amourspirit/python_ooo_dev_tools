from __future__ import annotations
from typing import Any, Tuple, TYPE_CHECKING, cast
import uno
from ..utils import props as mProps

from ..events.lo_events import Events
from ..events.props_named_event import PropsNamedEvent
from ..events.args.key_val_cancel_args import KeyValCancelArgs
from ..events.args.key_val_args import KeyValArgs
from abc import ABC, abstractmethod

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue


class StyleBase(ABC):
    def __init__(self, **kwargs) -> None:
        self._dv = {}
        for (key, value) in kwargs.items():
            if not value is None:
                self._dv[key] = value
        self._events = Events(source=self)
        self._events.on(PropsNamedEvent.PROP_SETTING, _on_props_setting)
        self._events.on(PropsNamedEvent.PROP_SET, _on_props_set)

    def _get(self, key: str) -> Any:
        return self._dv.get(key, None)

    def _set(self, key: str, val: Any) -> bool:
        if val is None:
            if key in self._dv:
                del self._dv[key]
            return False
        self._dv[key] = val
        return True

    def _clear(self) -> None:
        self._dv.clear()

    def _has(self, key: str) -> bool:
        return key in self._dv

    def _remove(self, key: str) -> bool:
        if self._has(key):
            del self._dv[key]
            return True
        return False

    def _is_supported(self, obj: object) -> bool:
        # can be used in child classe to for something like if mInfo.Info.support_service(obj, "com.sun.star.table.CellProperties"):
        raise NotImplemented

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attribures
        """
        # get current keys in internal dictionary
        return tuple(self._dv.keys())

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Oject that styles are to be applied.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        if len(self._dv) > 0:
            mProps.Props.set(obj, **self._dv)

    def on_property_setting(self, event_args: KeyValCancelArgs):
        """
        Raise for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_set(self, event_args: KeyValArgs):
        """
        Raise for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def get_props(self) -> Tuple[PropertyValue, ...]:
        """
        Gets instance properties

        Returns:
            Tuple[PropertyValue, ...]: Tuple of properties.
        """
        # see: setPropertyValues([in] sequence< com::sun::star::beans::PropertyValue > aProps)
        # https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertyAccess.html#a5ac97dfa6d796f4c794e2350e9130692
        if len(self._dv) == 0:
            return ()
        return mProps.Props.make_props(**self._dv)

    @property
    def has_attribs(self) -> bool:
        """Gets If instantance has any attributes set."""
        return len(self._dv)


def _on_props_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_setting(event_args)


def _on_props_set(source: Any, event_args: KeyValArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_set(event_args)


__all__ = ("StyleBase",)
