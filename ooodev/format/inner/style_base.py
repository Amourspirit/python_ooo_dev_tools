# region Imports
from __future__ import annotations
from typing import Any, Dict, NamedTuple, Tuple, TYPE_CHECKING, Type, TypeVar, cast, overload
import uno
from com.sun.star.container import XNameContainer
from com.sun.star.beans import XPropertySet

# import random
# import string

from ooodev.utils import props as mProps
from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.events.lo_events import Events
from ooodev.events.props_named_event import PropsNamedEvent
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs as KeyValCancelArgs
from ooodev.events.args.key_val_args import KeyValArgs as KeyValArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs as CancelEventArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventCallback
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.events.format_named_event import FormatNamedEvent as FormatNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.prop_pair import PropPair

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.style import CellStyle

# endregion Imports

# region Type Vars
TStyleBase = TypeVar(name="TStyleBase", bound="StyleBase")
TStyleMulti = TypeVar(name="TStyleMulti", bound="StyleMulti")
TStyleName = TypeVar(name="TStyleName", bound="StyleName")
_TStyleModifyMulti = TypeVar(name="_TStyleModifyMulti", bound="StyleModifyMulti")
# endregion Type Vars


# region Meta
class MetaStyle(type):
    def __call__(cls, *args, **kw):
        custom_args = kw.pop("_cattribs", None)
        obj = cls.__new__(cls, *args, **kw)
        # uniquie_id = "".join(random.choices(string.ascii_uppercase + string.digits, k=12))
        # object.__setattr__(obj, "_uniquie_id", uniquie_id)
        _events = Events(source=obj)
        object.__setattr__(obj, "_internal_events", _events)

        if custom_args:
            for key, value in cast(Dict[str, Any], custom_args).items():
                object.__setattr__(obj, key, value)
        obj.__init__(*args, **kw)
        return obj


# endregion Meta


# region Style Base Class
class StyleBase(metaclass=MetaStyle):
    """
    Base Styles class

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(self, **kwargs) -> None:
        # this property is used in child classes that have default instances
        # self._events = Events(source=self)

        # self._uniquie_id = "".join(random.choices(string.ascii_uppercase + string.digits, k=12))

        self._is_default_inst = False
        self._prop_parent = None

        self._dv = {}
        self._dv_bak = None

        for key, value in kwargs.items():
            if key.startswith("__"):
                # internal and not consider a property
                continue
            if not value is None:
                self._dv[key] = value
        super().__init__()
        self._set_style_internal_events()

    def _set_style_internal_events(self):
        def on_getting_cattribs(source: Any, event_args: CancelEventArgs) -> None:
            self._on_getting_cattribs(source=source, event_args=event_args)

        def on_clearing(source: Any, event_args: CancelEventArgs) -> None:
            self._on_clearing(source=source, event_args=event_args)

        def on_removing(source: Any, event_args: CancelEventArgs) -> None:
            self._on_removing(source=source, event_args=event_args)

        def on_setting(source: Any, event_args: KeyValCancelArgs) -> None:
            self._on_setting(source=source, event_args=event_args)

        def on_copying(source: Any, event_args: CancelEventArgs) -> None:
            self._on_copying(source=source, event_args=event_args)

        def on_backing_up(source: Any, event_args: KeyValCancelArgs) -> None:
            self.on_property_backing_up(source=source, event_args=event_args)

        def on_backed_up(source: Any, event_args: KeyValArgs) -> None:
            self.on_property_backed_up(source=source, event_args=event_args)

        def on_applying(source: Any, event_args: CancelEventArgs) -> None:
            self.on_applying(source=source, event_args=event_args)

        def on_applied(source: Any, event_args: EventArgs) -> None:
            self.on_applied(source=source, event_args=event_args)

        self._fn_on_getting_cattribs = on_getting_cattribs
        self._fn_on_clearing = on_clearing
        self._fn_on_removing = on_removing
        self._fn_on_setting = on_setting
        self._fn_on_copying = on_copying
        self._fn_on_backing_up = on_backing_up
        self._fn_on_backed_up = on_backed_up
        self._fn_on_applying = on_applying
        self._fn_on_applied = on_applied

        self._events.on("internal_cattribs", on_getting_cattribs)
        self._events.on(FormatNamedEvent.STYLE_CLEARING, on_clearing)
        self._events.on(FormatNamedEvent.STYLE_REMOVING, on_removing)
        self._events.on(FormatNamedEvent.STYLE_SETTING, on_setting)
        self._events.on(FormatNamedEvent.STYLE_COPYING, on_copying)
        self._events.on(FormatNamedEvent.STYLE_BACKING_UP, on_backing_up)
        self._events.on(FormatNamedEvent.STYLE_BACKED_UP, on_backed_up)
        self._events.on(FormatNamedEvent.STYLE_APPLYING, on_applying)
        self._events.on(FormatNamedEvent.STYLE_APPLIED, on_applied)

    # endregion Init

    # region Events

    def add_event_listener(self, event_name: str, callback: EventCallback) -> None:
        """
        Add an event listener to current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:

        See Also:
            - :py:class:`~.format_named_event.FormatNamedEvent`
            - :py:meth:`.remove_event_listener`

        Note:
            This method is generally only used in the context of child classes.
        """
        self._events.on(event_name, callback)

    def remove_event_listener(self, event_name: str, callback: EventCallback) -> None:
        """
        Remove an event listener from current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:

        See Also:
            - :py:class:`~.format_named_event.FormatNamedEvent`
            - :py:meth:`.add_event_listener`

        Note:
            This method is generally only used in the context of child classes.
        """
        self._events.remove(event_name, callback)

    # endregion Events

    # region style property methods

    def _get_properties(self) -> Dict[str, Any]:
        """Gets Key value pairs for the instance."""
        return self._dv

    def _get(self, key: str) -> Any:
        """Gets the property value"""
        return self._get_properties().get(key, None)

    def _set(self, key: str, val: Any) -> bool:
        """Sets a property value"""
        kvargs = KeyValCancelArgs("style_base", key=key, value=val)
        cargs = CancelEventArgs.from_args(kvargs)
        self._events.trigger(FormatNamedEvent.STYLE_SETTING, kvargs)
        self._events.trigger(FormatNamedEvent.STYLE_MODIFYING, cargs)
        if kvargs.cancel:
            return False
        if cargs.cancel:
            return False
        dv = self._get_properties()
        dv[kvargs.key] = kvargs.value
        self._events.trigger(FormatNamedEvent.STYLE_SET, KeyValArgs.from_args(kvargs))
        return True

    def _clear(self) -> None:
        """Clears all properties"""
        cargs = CancelEventArgs("style_base")
        self._events.trigger(FormatNamedEvent.STYLE_MODIFYING, cargs)
        self._events.trigger(FormatNamedEvent.STYLE_CLEARING, cargs)
        if cargs.cancel:
            return
        dv = self._get_properties()
        dv.clear()

    def _has(self, key: str) -> bool:
        """Gets if a property exist"""
        return key in self._get_properties()

    def _remove(self, key: str) -> bool:
        """Removes a property if it exist"""
        cargs = CancelEventArgs("style_base")
        cargs.event_data = key
        self._events.trigger(FormatNamedEvent.STYLE_REMOVING, cargs)
        self._events.trigger(FormatNamedEvent.STYLE_MODIFYING, cargs)
        if cargs.cancel:
            return
        if self._has(key):
            dv = self._get_properties()
            del dv[key]
            return True
        return False

    def _del_attribs(self, *attribs: str) -> bool:
        """Delete Attributes from instance if exist. Calls ``_on_deleting_attrib()``"""
        for attrib in attribs:
            if hasattr(self.__class__, attrib):
                kvargs = KeyValCancelArgs("style_base", key=attrib, value=getattr(self, attrib, None))
                self._on_deleting_attrib(self, kvargs)
                if kvargs.cancel:
                    continue
                delattr(self.__class_, attrib)
        for attrib in attribs:
            if hasattr(self, attrib):
                kvargs = KeyValCancelArgs("style_base", key=attrib, value=getattr(self, attrib, None))
                self._on_deleting_attrib(self, kvargs)
                if kvargs.cancel:
                    continue
                delattr(self, attrib)

    def _update(self, value: Dict[str, Any] | StyleBase) -> None:
        """Updates properties"""
        cargs = CancelEventArgs("style_base")
        cargs.event_data: Dict[str, Any] | StyleBase = value
        self._events.trigger(FormatNamedEvent.STYLE_UPDATING, cargs)
        self._events.trigger(FormatNamedEvent.STYLE_MODIFYING, cargs)
        if cargs.cancel:
            return
        dv = self._get_properties()
        if isinstance(cargs.event_data, StyleBase):
            dv.update(cargs.event_data._dv)
            return
        dv.update(cargs.event_data)

    # endregion style property methods

    # region Services
    def support_service(self, *service: str) -> bool:
        """
        Gets if service is supported.

        Args:
            service: expandable list of service names of UNO services such as ``com.sun.star.text.TextFrame``.

        Returns:
            bool: ``True`` if service is supported; Otherwise, ``Fasle``.
        """
        services = self._supported_services()
        for s in service:
            if s in services:
                return True
        return False

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
        return self._is_obj_service(obj)

    def _is_obj_service(self, obj: object) -> bool:
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

    def _print_not_valid_srv(self, method_name: str = ""):
        """
        Prints via ``Lo.print()`` notice that required service is missing

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
        srv = "service" if rs_len == 1 else "services"
        mLo.Lo.print(f"{self.__class__.__name__}{name}: object must support {srv}: {services}")

    # endregion Services

    # region Internal Methods
    def _props_set(self, obj: object, **kwargs: Any) -> None:
        # set properties. Can be overriden in child classes
        # may be usful to wrap in try statements in child classes
        mProps.Props.set(obj, **kwargs)

    def _copy_missing_attribs(self, src: TStyleBase, dst: TStyleBase, *args: str) -> None:
        """
        Copies attribs from source to dst if dst does not already have the attrib.

        Args:
            src (TStyleBase): Source
            dst (TStyleBase): Destination

        Returns:
            None:
        """
        for arg in args:
            if not hasattr(dst, arg) and hasattr(src, arg):
                setattr(dst, arg, getattr(src, arg))

    # region _props methods
    def _get_internal_cattribs(self) -> dict:
        cattribs = {
            "_props_internal_attributes": self._props,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }
        cargs = CancelEventArgs(self)
        cargs.event_data = cattribs
        self._events.trigger("internal_cattribs", cargs)
        if cargs.cancel:
            return None
        return cargs.event_data

    # endregion _props methods

    # endregion Internal Methods

    # region Methods
    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attributes
        """
        # get current keys in internal dictionary
        return tuple(self._get_properties().keys())

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Object that styles are to be applied.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Keyword Arguments:
            override_dv (Dic[str, Any], optional): if passed in this dictionary is used to set properties instead of internal dictionary of property values.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if "override_dv" in kwargs:
            dv = kwargs["override_dv"]
        else:
            dv = self._get_properties()
        if len(dv) > 0:
            if self._is_valid_obj(obj):
                cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
                cargs.event_data = self
                self._events.trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
                if cargs.cancel:
                    return
                events = Events(source=self)
                events.on(PropsNamedEvent.PROP_SETTING, _on_props_setting)
                events.on(PropsNamedEvent.PROP_SET, _on_props_set)
                # mProps.Props.set(obj, **dv)
                self._props_set(obj, **dv)
                events = None
                eargs = EventArgs.from_args(cargs)
                self._events.trigger(FormatNamedEvent.STYLE_APPLIED, eargs)
            else:
                self._print_not_valid_srv("apply")

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

    # region Copy()
    @overload
    def copy(self: TStyleBase) -> TStyleBase:
        ...

    @overload
    def copy(self: TStyleBase, **kwargs) -> TStyleBase:
        ...

    def copy(self: TStyleBase, **kwargs) -> TStyleBase:
        """Gets a copy of instance as a new instance"""
        cargs = CancelEventArgs(self)
        self._events.trigger(FormatNamedEvent.STYLE_COPYING, cargs)
        if cargs.cancel:
            if cargs.handled:
                return cargs.event_data
            else:
                mEx.CancelEventError(cargs)
        nu = self.__class__(**kwargs)
        nu._prop_parent = self._prop_parent
        # depending on python 3.7 builtin dictionary ordering
        dv = self._get_properties()
        if dv:
            # it is possible that that a new instance will have different property names thne the current instance.
            # This can happen because this class inherits from MetaStyle.
            # if ne contains a _props attribute (tuple of prop names) then use them to remap keys.
            # For instance a key of BorderLength may become ParaBoderLength.

            key_map = None
            p_len = len(nu._props)
            if p_len > 0 and p_len == len(self._props):
                key_map = {}
                for i, p_val in enumerate(self._props):
                    if p_val == "":
                        # some prop value may not be used in which case they are empty strings.
                        continue
                    if isinstance(p_val, PropPair):
                        nu_pair = cast(PropPair, nu._props[i])
                        if p_val.first:
                            key_map[p_val.first] = nu_pair.first
                        if p_val.second:
                            key_map[p_val.second] = nu_pair.second
                    else:
                        key_map[p_val] = nu._props[i]

            if key_map:
                for key, nu_val in key_map.items():
                    if self._has(key):
                        nu._set(nu_val, self._get(key))
            else:
                nu._update(self._get_properties())
        return nu

    # endregion Copy()

    # endregion Methods

    # region Backup/Restore

    def backup(self, obj: object) -> None:
        """
        Backs up Attributes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (object): Object to backup properties from.

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_BACKING_UP` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_BACKED_UP` :eventref:`src-docs-key-event`

        See Also:
            :py:meth:`~.style_base.StyleBase.restore`
        """
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv("Backup")
            return
        if self._dv_bak is None:
            self._dv_bak = {}
        for attr in self.get_attrs():
            val = mProps.Props.get(obj, attr, None)
            cargs = KeyValCancelArgs("style_base", key=attr, value=val)
            self._events.trigger(FormatNamedEvent.STYLE_BACKING_UP, cargs)
            if cargs.cancel:
                continue
            self._dv_bak[attr] = val
            eargs = KeyValArgs.from_args(cargs)
            self._events.trigger(FormatNamedEvent.STYLE_BACKED_UP, eargs)

    def restore(self, obj: object, clear: bool = False) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (object): Object to restore properties on.
            clear (bool): Determines if backup is cleared after restore. Default ``False``

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_PROPERTY_RESTORING` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_PROPERTY_RESTORED` :eventref:`src-docs-key-event`

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

    def _on_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        # called by _set()
        pass

    def _on_removing(self, source: Any, event_args: CancelEventArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        # called by _remove()
        pass

    def _on_getting_cattribs(self, source: Any, event_args: CancelEventArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        # called by _remove()
        pass

    def _on_clearing(self, source: Any, event_args: CancelEventArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        # called by _clear()
        pass

    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        # called by _set(), _remove(), _clear(), _update()
        pass

    def _on_copying(self, source: Any, event_args: CancelEventArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        pass

    def _on_deleting_attrib(self, source: Any, event_args: KeyValCancelArgs) -> None:
        # can be overridden in child classes to manage or modify setting
        pass

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_set(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_backing_up(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be backup up during backup

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args.
        """
        # can be overriden in child classes.
        pass

    def on_property_backed_up(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers before each property that is about to be set during backup.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_property_restore_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be set during restore

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overriden in child classes.
        event_args.set("on_property_restore_setting", True)
        self.on_property_setting(source, event_args)

    def on_property_restore_set(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that has been set during restore

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        event_args.set("on_property_restore_set", True)
        self.on_property_set(source, event_args)

    def on_applying(self, source: Any, event_args: CancelEventArgs) -> None:
        """
        Triggers before style/format is applied.

        Args:
            source (Any): Event Source.
            event_args (CancelEventArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    def on_applied(self, source: Any, event_args: EventArgs) -> None:
        """
        Triggers after style/format is applied.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overriden in child classes.
        pass

    # endregion Event Methods

    # region Dunder Methods

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, StyleBase):
            result = False
            try:
                for k, v in self._get_properties().items():
                    if oth._get(k) != v:
                        return False
                result = True
            except Exception:
                return False
            return result
        return NotImplemented

    # endregion Dunder Methods

    # region Named Container Methods

    def _container_get_service_name(self) -> str:
        raise NotImplementedError

    def _container_get_inst(self) -> XNameContainer:
        container = mLo.Lo.create_instance_msf(XNameContainer, self._container_get_service_name(), raise_err=True)
        return container

    def _container_add_value(
        self, name: str, obj: object, allow_update: bool = True, nc: XNameContainer | None = None
    ) -> None:
        if nc is None:
            nc = self._container_get_inst()
        if nc.hasByName(name):
            if allow_update:
                nc.replaceByName(name, obj)
                return
        else:
            nc.insertByName(name, obj)

    def _container_get_unique_el_name(self, prefix: str, nc: XNameContainer | None = None) -> str:
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
            nc = self._container_get_inst()
        names = nc.getElementNames()
        i = 1
        name = f"{prefix}{i}"
        while name in names:
            i += 1
            name = f"{prefix}{i}"
        return name

    def _container_get_value(self, name: str, nc: XNameContainer | None = None) -> Any:
        if not name:
            raise ValueError("Name is empty value. Expected a string name.")
        if nc is None:
            nc = self._container_get_inst()
        if nc.hasByName(name):
            return nc.getByName(name)
        return None

    # endregion Named Container Methods

    # region Properties

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instance has any attributes set."""
        return len(self._dv) > 0

    @property
    def prop_has_backup(self) -> bool:
        """Gets If instance has backup data set."""
        if self._dv_bak is None:
            return False
        return len(self._dv_bak) > 0

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.UNKNOWN

    @property
    def prop_parent(self) -> StyleBase | None:
        """Gets Parent Class"""
        return self._prop_parent

    @property
    def _props(self) -> Tuple[str, ...]:
        # placeholder for child classes. Usd in copy method.
        return ()

    @property
    def _events(self) -> Events:
        return self._internal_events

    # endregion Properties


# endregion Style Base Class


# region Module Internal Helper Classes
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


class _StyleInfo(NamedTuple):
    style: StyleBase
    args: _StyleMultArgs | None


# endregion Module Internal Helper Classes


# region Style Multi Class
class StyleMulti(StyleBase):
    """
    Multi style class.

    Supports appending styles via ``_append_style()`` (protected) method.
    When ``apply_style()`` is call all internal style instances are also applied.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(self, **kwargs) -> None:
        self._styles: Dict[str, _StyleInfo] = {}
        self._all_attributes = True
        super().__init__(**kwargs)

    # endregion Init

    # region Overrides

    def _set_style_internal_events(self) -> None:
        super()._set_style_internal_events()

        def on_multi_style_setting(source: Any, event_args: KeyValCancelArgs) -> None:
            self._on_multi_style_setting(source, event_args)

        def on_multi_style_set(source: Any, event_args: KeyValArgs) -> None:
            self._on_multi_style_set(source, event_args)

        def on_multi_style_removing(source: Any, event_args: CancelEventArgs) -> None:
            self._on_multi_style_removing(source, event_args)

        def on_multi_style_removed(source: Any, event_args: EventArgs) -> None:
            self._on_multi_style_removed(source, event_args)

        def on_multi_style_updating(source: Any, event_args: CancelEventArgs) -> None:
            self._on_multi_style_updating(source, event_args)

        def on_multi_style_updated(source: Any, event_args: EventArgs) -> None:
            self._on_multi_style_updated(source, event_args)

        def on_multi_child_style_applying(source: Any, event_args: KeyValCancelArgs) -> None:
            self._on_multi_child_style_applying(source, event_args)

        def on_multi_child_style_applied(source: Any, event_args: KeyValArgs) -> None:
            self._on_multi_child_style_applied(source, event_args)

        self._fn_on_multi_style_setting = on_multi_style_setting
        self._fn_on_multi_style_set = on_multi_style_set
        self._fn_on_multi_style_removing = on_multi_style_removing
        self._fn_on_multi_style_removed = on_multi_style_removed
        self._fn_on_multi_style_updating = on_multi_style_updating
        self._fn_on_multi_style_updated = on_multi_style_updated
        self._fn_on_multi_child_style_applying = on_multi_child_style_applying
        self._fn_on_multi_child_style_applied = on_multi_child_style_applied

        self._events.on(FormatNamedEvent.MULTI_STYLE_SETTING, on_multi_style_setting)
        self._events.on(FormatNamedEvent.MULTI_STYLE_SET, on_multi_style_set)
        self._events.on(FormatNamedEvent.MULTI_STYLE_REMOVING, on_multi_style_removing)
        self._events.on(FormatNamedEvent.MULTI_STYLE_REMOVED, on_multi_style_removed)
        self._events.on(FormatNamedEvent.MULTI_STYLE_UPDATING, on_multi_style_updating)
        self._events.on(FormatNamedEvent.MULTI_STYLE_UPDATED, on_multi_style_updated)
        self._events.on(FormatNamedEvent.STYLE_MULTI_CHILD_APPLYING, on_multi_child_style_applying)
        self._events.on(FormatNamedEvent.STYLE_MULTI_CHILD_APPLIED, on_multi_child_style_applied)

    # region apply()

    def _apply_direct(self, obj: object, **kwargs) -> None:
        """Calls super apply directly"""
        super().apply(obj, **kwargs)

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies style of current instance and all other internal style instances.

        Args:
            obj (object): UNO Object that styles are to be applied.
        """
        styles = self._get_multi_styles()
        for key, info in styles.items():
            kvargs = KeyValCancelArgs("StyleMulti.apply", key=key, value=info)
            self._events.trigger(FormatNamedEvent.STYLE_MULTI_CHILD_APPLYING, kvargs)
            if kvargs.cancel:
                continue
            style, kw = info
            if kw:
                style.apply(obj, **kw.kwargs)
            else:
                style.apply(obj)
            self._events.trigger(FormatNamedEvent.STYLE_MULTI_CHILD_APPLIED, KeyValArgs.from_args(kvargs))
        # apply this instance properties after all others styles.
        # allows this instance to overwrite properties set by multi styles if needed.
        super().apply(obj, **kwargs)

    # endregion apply()

    # region Copy()
    @overload
    def copy(self: TStyleMulti) -> TStyleMulti:
        ...

    @overload
    def copy(self: TStyleMulti, **kwargs) -> TStyleMulti:
        ...

    def copy(self: TStyleMulti, **kwargs) -> TStyleMulti:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        styles = self._get_multi_styles()
        for key, info in styles.items():
            style, kw = info
            if kw:
                cp._set_style(key, style.copy(**kwargs), **kw.kwargs)
            else:
                cp._set_style(key, style.copy(**kwargs))
        return cp

    # endregion Copy()

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

    # endregion Overrides

    # region Internal Methods

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
        kvargs = KeyValCancelArgs("_set_style", key=key, value=style)
        self._events.trigger(FormatNamedEvent.MULTI_STYLE_SETTING, kvargs)
        if kvargs.cancel:
            return
        styles = self._get_multi_styles()
        if style._prop_parent is None:
            style._prop_parent = self
        if len(attrs) + len(kwargs) == 0:
            styles[key] = _StyleInfo(style, None)
        else:
            styles[key] = _StyleInfo(style, _StyleMultArgs(*attrs, **kwargs))
        self._events.trigger(FormatNamedEvent.MULTI_STYLE_SET, KeyValArgs.from_args(kvargs))

    def _update_style(self, value: StyleMulti) -> None:
        cargs = CancelEventArgs("_update_style")
        cargs.event_data = value
        self._events.trigger(FormatNamedEvent.MULTI_STYLE_UPDATING, cargs)
        if cargs.cancel:
            return
        self._get_multi_styles().update(value._styles)
        self._events.trigger(FormatNamedEvent.MULTI_STYLE_UPDATED, EventArgs.from_args(cargs))

    def _remove_style(self, key: str) -> bool:
        cargs = CancelEventArgs("_set_style")
        cargs.event_data = key
        self._events.trigger(FormatNamedEvent.MULTI_STYLE_REMOVING, cargs)
        if cargs.cancel:
            return
        result = False
        styles = self._get_multi_styles()
        if key in styles:
            del styles[key]
            result = True
            # only trigger if removed
            self._events.trigger(FormatNamedEvent.MULTI_STYLE_REMOVED, EventArgs.from_args(cargs))
        return result

    def _get_style(self, key: str) -> _StyleInfo | None:
        return self._get_multi_styles().get(key, None)

    def _get_style_inst(self, key: str) -> StyleBase | None:
        style = self._get_style(key)
        if style is None:
            return None
        return style.style

    def _has_style(self, key: str) -> bool:
        return key in self._get_multi_styles()

    def _get_multi_styles(self) -> Dict[str, _StyleInfo]:
        return self._styles

    # endregion Internal Methods

    # region Methods
    def backup(self, obj: object) -> None:
        """
        Backs up Attributes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (object): Object to backup properties from.

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_BACKING_UP` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_BACKED_UP` :eventref:`src-docs-key-event`

        See Also:
            :py:meth:`~.style_base.StyleMulti.restore`
        """
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv("Backup")
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
            clear (bool): Determines if backup is cleared after restore. Default ``False``

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_PROPERTY_RESTORING` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_PROPERTY_RESTORED` :eventref:`src-docs-key-event`

        See Also:
            :py:meth:`~.style_base.StyleMulti.backup`
        """
        super().restore(obj=obj, clear=clear)
        styles = self._get_multi_styles()
        for _, info in styles.items():
            style, _ = info
            style.restore(obj=obj, clear=clear)

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attributes
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

    # endregion Methods

    # region event handlers
    def _on_multi_style_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        pass

    def _on_multi_style_set(self, source: Any, event_args: KeyValArgs) -> None:
        pass

    def _on_multi_style_removing(self, source: Any, event_args: CancelEventArgs) -> None:
        pass

    def _on_multi_style_removed(self, source: Any, event_args: EventArgs) -> None:
        pass

    def _on_multi_style_updating(self, source: Any, event_args: CancelEventArgs) -> None:
        pass

    def _on_multi_style_updated(self, source: Any, event_args: EventArgs) -> None:
        pass

    def _on_multi_child_style_applying(self, source: Any, event_args: KeyValCancelArgs) -> None:
        pass

    def _on_multi_child_style_applied(self, source: Any, event_args: KeyValArgs) -> None:
        pass

    # endregion event handlers

    # region Properties

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instance has any attributes set."""
        return len(self._dv) + len(self._styles) > 0

    @property
    def prop_has_backup(self) -> bool:
        """Gets If instance or any added style has backup data set."""
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

    # endregion Properties


# endregion Style Multi Class

# region Style Modify Multi class


class StyleModifyMulti(StyleMulti):
    """
    Para Style Base

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(self, **kwargs) -> None:
        super().__init__(**kwargs)
        self._style_family_name = ""

    # endregion Init

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.Style",
            "com.sun.star.style.ParagraphStyle",
            "com.sun.star.beans.PropertySet",
        )

    def _is_valid_obj(self, obj: object) -> bool:
        if mLo.Lo.is_uno_interfaces(obj, "com.sun.star.style.XStyle"):
            return self._is_obj_service(obj)
        else:
            valid = self._is_obj_service(obj)
            if valid:
                return True
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # region apply()

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO Object such as a Document, Spreadsheet, or ``XStyle``

        Returns:
            None:
        """

        if mLo.Lo.is_uno_interfaces(obj, "com.sun.star.style.XStyle"):
            if not self._is_obj_service(obj):
                self._print_not_valid_srv(method_name="apply")
                return
            p = mLo.Lo.qi(XPropertySet, obj)
            if p is None:
                mLo.Lo.print(
                    f"{self.__class__.__name__}.apply(): Not a UNO Object for style. Unable to set Style Properties"
                )
                return
        else:

            if not self._is_valid_doc(obj):
                mLo.Lo.print(
                    f"{self.__class__.__name__}.apply(): Not a UNO Object for style. Unable to set Style Properties"
                )
                return
            p = self.get_style_props(obj)
        super().apply(p, **kwargs)

    # endregion apply()

    # region copy()
    @overload
    def copy(self: _TStyleModifyMulti) -> _TStyleModifyMulti:
        ...

    @overload
    def copy(self: _TStyleModifyMulti, **kwargs) -> _TStyleModifyMulti:
        ...

    def copy(self: _TStyleModifyMulti, **kwargs) -> _TStyleModifyMulti:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp.prop_style_name = self.prop_style_name
        return cp

    # endregion copy()

    # endregion Overrides

    # region internal methods

    def _is_valid_doc(self, obj: object) -> bool:
        return True
        # return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)

    def _get_style_family_name(self) -> str:
        if not self._style_family_name:
            raise ValueError("Must set internal property _style_family_name")
        return self._style_family_name

    # endregion internal methods

    # region Methods

    def get_style_props(self, doc: object) -> XPropertySet:
        """
        Gets the Style Properties

        Args:
            doc (object): UNO Document Object.

        Raises:
            NotSupportedDocumentError: If document is not supported.

        Returns:
            XPropertySet: Styles properties property set.
        """
        if not self._is_valid_doc(doc):
            raise mEx.NotSupportedDocumentError
        return mInfo.Info.get_style_props(doc, self._get_style_family_name(), self.prop_style_name)

    # endregion Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_style_name(self) -> str:
        """
        Gets/Sets property Style Name.

        Raises:
            NotImplementedError:
        """
        raise NotImplementedError

    @prop_style_name.setter
    def prop_style_name(self, value: str):
        raise NotImplementedError

    @property
    def prop_style_family_name(self) -> str:
        """Gets/Set Style Family Name"""
        return self._style_family_name

    @prop_style_family_name.setter
    def prop_style_family_name(self, value: str):
        self._style_family_name = value

    # endregion Properties


# endregion Style Modify Multi class


# region StyleName Class
class StyleName(StyleBase):
    """Style Name Base Class"""

    # region Init
    def __init__(self, name: Any, **kwargs) -> None:
        """
        Constructor

        Args:
            name (Any): Style Name.

        Raises:
            ValueError: If Name is ``None`` or empty string.
        """
        if not name and name != "":
            raise ValueError("Name is required.")
        init_vars = {self._get_property_name(): str(name)}
        init_vars.update(kwargs)
        super().__init__(**init_vars)

    # endregion Init

    # region methods
    def get_style_props(self) -> XPropertySet:
        """
        Gets Style as ``XPropertySet`` that contains all style properties.

        Returns:
            XPropertySet: Returns result also implements ``com.sun.star.style.XStyle``

        .. versionadded:: 0.9.2
        """
        props = mInfo.Info.get_style_props(
            doc=mLo.Lo.this_component, family_style_name=self._get_family_style_name(), prop_set_nm=self.prop_name
        )
        return props

    # endregion methods

    # region internal methods
    def _get_family_style_name(self) -> str:
        raise NotImplementedError

    def _get_property_name(self) -> str:
        try:
            return self._style_property_name
        except AttributeError:
            raise NotImplementedError

    # endregion internal methods

    # region Overrides
    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # endregion Overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[TStyleName], obj: object) -> TStyleName:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[TStyleName], obj: object, **kwargs) -> TStyleName:
        ...

    @classmethod
    def from_obj(cls: Type[TStyleName], obj: object, **kwargs) -> TStyleName:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TStyleName: ``TStyleName`` instance that represents ``obj`` style.
        """
        inst = cls(**kwargs)

        if mInfo.Info.support_service(obj, "com.sun.star.style.CellStyle"):
            cs = cast("CellStyle", obj)
            pname = cs.getName()
            inst.prop_name = pname
            return inst

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        pname = mProps.Props.get(obj, inst._get_property_name(), "")
        if pname:
            inst.prop_name = pname
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.STATIC
        return self._format_kind_prop

    @property
    def prop_name(self) -> str:
        """Gets/Sets style name"""
        return self._get(self._get_property_name())

    @prop_name.setter
    def prop_name(self, value: Any) -> None:
        if not value:
            raise ValueError("Value must not be None or Empty string.")
        self._set(self._get_property_name(), str(value))

    # endregion Properties


# endregion StyleName Class


# region Props Property event handlers
def _on_props_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_setting(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_APPLYING, event_args)


def _on_props_set(source: Any, event_args: KeyValArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_set(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_APPLIED, event_args)


def _on_props_restore_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_restore_setting(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_RESTORING, event_args)


def _on_props_restore_set(source: Any, event_args: KeyValArgs, *args, **kwargs) -> None:
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_restore_set(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_RESTORED, event_args)


# endregion Props Property event handlers

__all__ = ("StyleBase", "StyleMulti", "StyleModifyMulti", "StyleName", "TStyleBase", "TStyleMulti", "TStyleName")
