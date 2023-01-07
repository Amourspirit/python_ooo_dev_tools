from __future__ import annotations
from typing import Any, Dict, Tuple, TYPE_CHECKING, cast
import uno
from ..utils import props as mProps

from ..events.lo_events import Events
from ..events.props_named_event import PropsNamedEvent
from ..events.args.key_val_cancel_args import KeyValCancelArgs
from ..events.args.key_val_args import KeyValArgs
from ..utils.type_var import T
from abc import ABC

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue


class StyleBase(ABC):
    """
    Base Styles class

    .. versionadded:: 0.9.0
    """

    def __init__(self, **kwargs) -> None:
        self._dv = {}

        for (key, value) in kwargs.items():
            if not value is None:
                self._dv[key] = value

    def _get(self, key: str) -> Any:
        return self._dv.get(key, None)

    def _set(self, key: str, val: Any) -> bool:
        cargs = KeyValCancelArgs("style_base", key=key, value=val)
        self._on_setting(cargs)
        if cargs.cancel:
            return False
        self._dv[cargs.key] = cargs.value
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

    def _update(self, kv: Dict[str, Any]) -> None:
        self._dv.update(kv)

    def _on_setting(self, event: KeyValCancelArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        # called by _set()
        pass

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
        if self.has_attribs:
            events = Events(source=self)
            events.on(PropsNamedEvent.PROP_SETTING, _on_props_setting)
            events.on(PropsNamedEvent.PROP_SET, _on_props_set)
            mProps.Props.set(obj, **self._dv)
            events = None

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
        return len(self._dv) > 0

    def copy(self: T) -> T:
        nu = super(StyleBase, self.__class__).__new__(self.__class__)
        nu.__init__()
        nu._update(self._dv)
        return nu


class StyleMulti(StyleBase):
    """
    Multi style class.

    Supports appending styles via ``_append_style()`` (protected) method.
    When ``apply_style()`` is call all internal style instances are also applied.

    .. versionadded:: 0.9.0
    """

    def __init__(self, **kwargs) -> None:
        super().__init__(**kwargs)
        self._styles: Dict[int, StyleBase] = {}

    def _set_style(self, key: str, style: StyleBase) -> None:
        self._styles[key] = style

    def _remove_style(self, key: str) -> bool:
        if key in self._styles:
            del self._styles[key]
            return True
        return False

    def _get_style(self, key: str) -> StyleBase | None:
        return self._styles.get(key, None)

    def _has_style(self, key: str) -> bool:
        return key in self._styles

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies style of current instance and all other internal style instances.

        Args:
            obj (object): UNO Oject that styles are to be applied.
        """
        super().apply_style(obj, **kwargs)
        for _, style in self._styles.items():
            style.apply_style(obj, **kwargs)

    def copy(self: T) -> T:
        cp = super().copy()
        for key, style in self._styles.items():
            cp._set_style(key, style.copy())
        return cp


def _on_props_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_setting(event_args)


def _on_props_set(source: Any, event_args: KeyValArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_set(event_args)


__all__ = ("StyleBase", "StyleMulti")
