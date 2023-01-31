from __future__ import annotations
from typing import Any, Dict, Tuple, TYPE_CHECKING, cast
import uno

from ..utils import props as mProps
from ..utils import info as mInfo
from ..utils import lo as mLo
from ..events.lo_events import Events
from ..events.props_named_event import PropsNamedEvent
from ..events.args.key_val_cancel_args import KeyValCancelArgs as KeyValCancelArgs
from ..events.args.key_val_args import KeyValArgs as KeyValArgs
from ..events.args.cancel_event_args import CancelEventArgs as CancelEventArgs
from ..events.args.event_args import EventArgs as EventArgs
from ..utils.type_var import T
from .kind.format_kind import FormatKind
from ..events.format_named_event import FormatNamedEvent as FormatNamedEvent
from ..events.event_singleton import _Events
from abc import ABC

from com.sun.star.container import XNameContainer

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue


class StyleBase(ABC):
    """
    Base Styles class

    .. versionadded:: 0.9.0
    """

    def __init__(self, **kwargs) -> None:
        self._dv = {}
        self._dv_bak = None

        for (key, value) in kwargs.items():
            if not value is None:
                self._dv[key] = value
        super().__init__()

    # region style property methods

    def _get_properties(self) -> Dict[str, Any]:
        """Gets Key value pairs for the instance."""
        return self._dv

    def _get(self, key: str) -> Any:
        """Gets the property value"""
        return self._dv.get(key, None)

    def _set(self, key: str, val: Any) -> bool:
        """Sets a property value"""
        cargs = KeyValCancelArgs("style_base", key=key, value=val)
        self._on_setting(cargs)
        if cargs.cancel:
            return False
        self._dv[cargs.key] = cargs.value
        return True

    def _clear(self) -> None:
        """Clears all properties"""
        self._dv.clear()

    def _has(self, key: str) -> bool:
        """Gets if a property exist"""
        return key in self._dv

    def _remove(self, key: str) -> bool:
        """Removes a property if it exist"""
        if self._has(key):
            del self._dv[key]
            return True
        return False

    def _update(self, value: Dict[str, Any] | StyleBase) -> None:
        """Updates properties"""
        if isinstance(value, StyleBase):
            self._dv.update(value._dv)
            return
        self._dv.update(value)

    # endregion style property methods

    # region Services
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of suported services for the style such as (``com.sun.star.style.ParagraphProperties``,)

        Raises:
            NotImplementedError: If not implemented in child class

        Returns:
            Tuple[str, ...]: Supported services
        """
        raise NotImplementedError

    def _is_valid_obj(self, obj: object) -> bool:
        """
        Gets if ``obj`` supports one of the services required by style class

        Args:
            obj (object): UNO object that must have requires service

        Returns:
            bool: ``True`` if has a required service; Otherwise, ``False``
        """
        rs = self._supported_services()
        if rs:
            return mInfo.Info.support_service(obj, *rs)
        # if style class has no required services then return True
        return True

    def _print_not_valid_obj(self, method_name: str = ""):
        """
        Prints via ``Lo.print()`` notice that requied service is missing

        Args:
            method_name (str, optional): Calling method name.
        """
        rs = self._supported_services()
        rs_len = len(rs)
        if method_name:
            name = f".{method_name}()"
        else:
            name = ""
        if rs_len == 0:
            mLo.Lo.print(f"{self.__class__.__name__}{name}: object is not valid.")
            return
        services = ", ".join(rs)
        srv = "service" if rs_len == 1 else "serivces"
        mLo.Lo.print(f"{self.__class__.__name__}{name}: object must support {srv}: {services}")

    # endregion Services

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attribures
        """
        # get current keys in internal dictionary
        return tuple(self._dv.keys())

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Oject that styles are to be applied.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if len(self._dv) > 0:
            if self._is_valid_obj(obj):
                cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
                cargs.event_data = self
                self.on_applying(cargs)
                if cargs.cancel:
                    return
                _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
                if cargs.cancel:
                    return
                events = Events(source=self)
                events.on(PropsNamedEvent.PROP_SETTING, _on_props_setting)
                events.on(PropsNamedEvent.PROP_SET, _on_props_set)
                # mProps.Props.set(obj, **self._dv)
                self._props_set(obj, **self._dv)
                events = None
                eargs = EventArgs.from_args(cargs)
                self.on_applied(eargs)
                _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)
            else:
                self._print_not_valid_obj("apply")

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        # set properties. Can be overriden in child classes
        # may be usful to wrap in try statements in child classes
        mProps.Props.set(obj, **kwargs)

    # region Backup/Restore

    def backup(self, obj: object) -> None:
        """
        Backs up Attriubes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (object): Object to backup properties from.

        Returns:
            None:

        See Also:
            :py:meth:`~.style_base.StyleBase.restore`
        """
        if not self._is_valid_obj(obj):
            self._print_not_valid_obj("Backup")
            return
        if self._dv_bak is None:
            self._dv_bak = {}
        for attr in self.get_attrs():
            val = mProps.Props.get(obj, attr, None)
            cargs = KeyValCancelArgs("style_base", key=attr, value=val)
            self.on_property_backing_up(cargs)
            if cargs.cancel:
                continue
            self._dv_bak[attr] = val
            eargs = KeyValArgs.from_args(cargs)
            self.on_property_backed_up(eargs)

    def restore(self, obj: object, clear: bool = False) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (object): Object to restore properties on.
            clear (bool): Determines if backup is cleared after resore. Default ``False``

        Returns:
            None:

        See Also:
            :py:meth:`~.style_base.StyleBase.backup`
        """
        if self._dv_bak is None:
            return
        if len(self._dv_bak) > 0:
            events = Events(source=self)
            events.on(PropsNamedEvent.PROP_SETTING, _on_props_restore_setting)
            events.on(PropsNamedEvent.PROP_SET, _on_props_restore_set)
            mProps.Props.set(obj, **self._dv_bak)
            events = None
            if clear:
                self._dv_bak.clear()

    # endregion Backup/Restore

    # region Event Methods

    def _on_setting(self, event: KeyValCancelArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        # called by _set()
        pass

    def on_property_setting(self, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_set(self, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_backing_up(self, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be backup up during backup

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_backed_up(self, event_args: KeyValArgs) -> None:
        """
        Triggers before each property that is about to be set during restore

        Args:
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_restore_setting(self, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be set during restore

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overriden in child classes.
        event_args.set("on_property_restore_setting", True)
        self.on_property_setting(event_args)

    def on_property_restore_set(self, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that has been set during restore

        Args:
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        event_args.set("on_property_restore_set", True)
        self.on_property_set(event_args)

    def on_applying(self, event_args: CancelEventArgs) -> None:
        """
        Triggers before style/format is applied.

        Args:
            event_args (CancelEventArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_applied(self, event_args: EventArgs) -> None:
        """
        Triggers after style/format is applied.

        Args:
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    # endregion Event Methods

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

    def copy(self: T) -> T:
        """Gets a copy of instance as a new instance"""
        nu = super(StyleBase, self.__class__).__new__(self.__class__)
        nu.__init__()
        nu._update(self._dv)
        return nu

    # region Dunder Methods

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, StyleBase):
            result = False
            try:
                for k, v in self._get_properties().items():
                    if oth._get(k) != v:
                        break
                result = True
            except Exception:
                return False
            return result
        return NotImplemented

    # endregion Dunder Methods

    # region Named Container Methods
    # endregion Named Container Methods

    def _get_container_service_name(self) -> str:
        raise NotImplementedError

    def _get_name_container(self) -> XNameContainer:
        container = mLo.Lo.create_instance_msf(XNameContainer, self._get_container_service_name(), raise_err=True)
        return container

    def _add_value_to_container(
        self, name: str, obj: object, allow_update: bool = True, nc: XNameContainer | None = None
    ) -> None:
        if nc is None:
            nc = self._get_name_container()
        if nc.hasByName(name):
            if allow_update:
                nc.replaceByName(name, obj)
                return
        else:
            nc.insertByName(name, obj)

    def _get_unnique_container_el_name(self, prefix: str, nc: XNameContainer | None = None) -> str:
        """
        Gets the next name that does not exist in the container.

        Lets say ``prefix`` is ``Transparency `` then names are search in sequence.
        ``Transparency 1``, ``Transparency 3``, ``Transparency 3``, etc until a unique name is found.

        Args:
            prefix (str): Any string such as ``Transparency ``
            nc (XNameContainer | None, optional): Container. Defaults to None.

        Returns:
            str: Unique name
        """
        if nc is None:
            nc = self._get_name_container()
        names = nc.getElementNames()
        i = 1
        name = f"{prefix}{i}"
        while name in names:
            i += 1
            name = f"{prefix}{i}"
        return name

    # region Properties

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instantance has any attributes set."""
        return len(self._dv) > 0

    @property
    def prop_has_backup(self) -> bool:
        """Gets If instantance has backup data set."""
        if self._dv_bak is None:
            return False
        return len(self._dv_bak) > 0

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.UNKNOWN

    # endregion Properties


class _StyleMultArgs:
    """Generic Args"""

    def __init__(self, *attrs, **kwargs):
        """
        Constructor
        """
        self._args = attrs[:]
        self._kwargs = kwargs.copy()

    @property
    def attrs(self) -> tuple:
        """
        Gets attribs tuple.

        This is a copy of ``args`` passed into constructor.
        """
        return self._args

    @property
    def kwargs(self) -> Dict[str, Any]:
        """
        Gets kwargs Dictionary

        This is a copy of ``kwargs`` passed into constructor
        """
        return self._kwargs


class StyleMulti(StyleBase):
    """
    Multi style class.

    Supports appending styles via ``_append_style()`` (protected) method.
    When ``apply_style()`` is call all internal style instances are also applied.

    .. versionadded:: 0.9.0
    """

    def __init__(self, **kwargs) -> None:
        super().__init__(**kwargs)
        self._styles: Dict[str, Tuple[StyleBase, _StyleMultArgs | None]] = {}
        self._all_attributes = True

    def _set_style(self, key: str, style: StyleBase, *attrs, **kwargs) -> None:
        """
        Sets style

        Args:
            key (str): key store style info
            style (StyleBase): style
            attrs: Exapandable list attributes that style sets.
                The values added here are added when get_attrs() method is called.
                This is used for backup and restore in Write Module.
            kwargs: Expandalble key value args to that are to be passed to style when ``apply_style()`` is called.
        """
        styles = self._get_multi_styles()
        if len(attrs) + len(kwargs) == 0:
            styles[key] = (style, None)
        else:
            styles[key] = (style, _StyleMultArgs(*attrs, **kwargs))

    def _update_style(self, value: StyleMulti) -> None:
        self._get_multi_styles().update(value._styles)

    def _remove_style(self, key: str) -> bool:
        styles = self._get_multi_styles()
        if key in styles:
            del styles[key]
            return True
        return False

    def _get_style(self, key: str) -> Tuple[StyleBase, _StyleMultArgs | None] | None:
        return self._get_multi_styles().get(key, None)

    def _has_style(self, key: str) -> bool:
        return key in self._get_multi_styles()

    def _get_multi_styles(self) -> Dict[str, Tuple[StyleBase, _StyleMultArgs | None]]:
        return self._styles

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instantance has any attributes set."""
        return len(self._dv) + len(self._styles) > 0

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies style of current instance and all other internal style instances.

        Args:
            obj (object): UNO Oject that styles are to be applied.
        """
        super().apply(obj, **kwargs)
        styles = self._get_multi_styles()
        for _, info in styles.items():
            style, kw = info
            if kw:
                style.apply(obj, **kw.kwargs)
            else:
                style.apply(obj)

    def copy(self: T) -> T:
        cp = super().copy()
        styles = self._get_multi_styles()
        for key, info in styles.items():
            style, kw = info
            if kw:
                cp._set_style(key, style.copy(), **kw.kwargs)
            else:
                cp._set_style(key, style.copy())
        return cp

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, StyleMulti):
            result = super().__eq__(oth)
            if result is False:
                return False
            styles = self._get_multi_styles()
            for key, info in styles.items():
                style, _ = info
                style_other = oth._get_style(key)
                if style_other is None:
                    result = False
                    break

                result = style == style_other[0]
                if result is False:
                    break
            return result
        return NotImplemented

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attribures
        """
        # get current keys in internal dictionary
        props = self._get_properties()
        attrs = set(props.keys())
        if self._all_attributes:
            styles = self._get_multi_styles()
            if styles:
                for _, info in styles.items():
                    _, args = info
                    if args:
                        attrs.update(args.attrs)
        return tuple(attrs)

    def backup(self, obj: object) -> None:
        """
        Backs up Attriubes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (object): Object to backup properties from.

        Returns:
            None:

        See Also:
            :py:meth:`~.style_base.StyleMulti.restore`
        """
        if not self._is_valid_obj(obj):
            self._print_not_valid_obj("Backup")
            return
        try:
            self._all_attributes = False
            super().backup(obj)
            styles = self._get_multi_styles()
            for _, info in styles.items():
                style, _ = info
                style.backup(obj)
        finally:
            self._all_attributes = True

    def restore(self, obj: object, clear: bool = False) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (object): Object to restore properties on.
            clear (bool): Determines if backup is cleared after resore. Default ``False``

        Returns:
            None:

        See Also:
            :py:meth:`~.style_base.StyleMulti.backup`
        """
        super().restore(obj=obj, clear=clear)
        styles = self._get_multi_styles()
        for _, info in styles.items():
            style, _ = info
            style.restore(obj=obj, clear=clear)

    @property
    def prop_has_backup(self) -> bool:
        """Gets If instantance or any added style has backup data set."""
        result = False
        styles = self._get_multi_styles()
        for _, info in styles.items():
            style, _ = info
            if style.prop_has_backup:
                result = True
                break
        if result:
            return True
        if self._dv_bak is None:
            return False
        return len(self._dv_bak) > 0


def _on_props_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_setting(event_args)


def _on_props_set(source: Any, event_args: KeyValArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_set(event_args)


def _on_props_restore_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_restore_setting(event_args)


def _on_props_restore_set(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_restore_set(event_args)


__all__ = ("StyleBase", "StyleMulti")
