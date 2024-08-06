"""Base Styles class"""

# pylint: disable=too-many-lines
# pylint: disable=broad-exception-raised
# pylint: disable=unused-import
# pylint: disable=useless-import-alias
# region Imports
from __future__ import annotations
from typing import Any, Dict, NamedTuple, Tuple, TYPE_CHECKING, Type, TypeVar, cast, overload
import contextlib
import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameContainer
from com.sun.star.lang import XMultiServiceFactory

# import random
# import string

from ooodev.utils import props as mProps
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.events.lo_events import Events, event_ctx
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
from ooodev.units.unit_obj import UnitT
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.style import CellStyle
    from ooodev.proto.event_observer import EventObserver
    from ooodev.format.proto.style_t import StyleT
else:
    PropertyValue = Any
    CellStyle = Any
    StyleT = Any

# endregion Imports

# region Type Vars
TStyleBase = TypeVar("TStyleBase", bound="StyleBase")  # pylint: disable=invalid-name
TStyleMulti = TypeVar("TStyleMulti", bound="StyleMulti")  # pylint: disable=invalid-name
TStyleName = TypeVar("TStyleName", bound="StyleName")  # pylint: disable=invalid-name
_TStyleModifyMulti = TypeVar("_TStyleModifyMulti", bound="StyleModifyMulti")  # pylint: disable=invalid-name
# endregion Type Vars


# region Meta
class MetaStyle(type):
    """Meta class for styles"""

    def __call__(cls, *args, **kw):
        # sourcery skip: instance-method-first-arg-name
        custom_args = kw.pop("_cattribs", None)
        obj = cls.__new__(cls, *args, **kw)  # type: ignore
        # unique_id = "".join(random.choices(string.ascii_uppercase + string.digits, k=12))
        # object.__setattr__(obj, "_unique_id", unique_id)
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

        # self._unique_id = "".join(random.choices(string.ascii_uppercase + string.digits, k=12))

        self._is_default_inst = False
        self._prop_parent: Any = None
        self.__update_obj = None

        self._dv = {}
        self._dv_bak = None

        for key, value in kwargs.items():
            if key.startswith("__"):
                # internal and not consider a property
                continue
            if value is not None:
                self._dv[key] = value
        super().__init__()
        self._set_style_internal_events()

    # region Update Methods
    def has_update_obj(self) -> bool:
        """
        Gets if the update object is set for the style instance.

        Returns:
            bool: ``True`` if the update object is set; Otherwise, ``False``.

        .. versionadded:: 0.27.0
        """
        return self.__update_obj is not None

    def get_update_obj(self) -> Any:
        """
        Gets the update object for the style instance.

        Returns:
            Any: The update object if set; Otherwise, ``None``.

        .. versionadded:: 0.28.2
        """
        return self.__update_obj

    def set_update_obj(self, obj: Any) -> None:
        """
        Sets the update object for the style instance.

        Args:
            obj (Any): Object used to apply style to when update is called.

        Returns:
            None:

        .. versionadded:: 0.27.0
        """
        # cannot set a weak ref to to a pyuno object.
        self.__update_obj = obj

    @overload
    def update(self) -> bool:
        """
        Applies the style to the update object.

        Returns:
            bool: Returns ``True`` if the update object is set and the style is applied; Otherwise, ``False``.
        """
        ...

    @overload
    def update(self, **kwargs: Any) -> bool:
        """
        Applies the style to the update object.

        Returns:
            bool: Returns ``True`` if the update object is set and the style is applied; Otherwise, ``False``.
            kwargs: Expandable list of key value pairs that may be used in child classes.
        """
        ...

    def update(self, **kwargs: Any) -> bool:
        """
        Applies the style to the update object.

        Returns:
            bool: Returns ``True`` if the update object is set and the style is applied; Otherwise, ``False``.
            kwargs: Expandable list of key value pairs that may be used in child classes.

        Note:
            If the ``apply()`` method has been called then usually the style will set the update object.
            For some style it may be necessary to set the update object manually using the ``set_update_obj()`` method.
            The ``has_update_obj()`` method can be used to get the update object as been set

        .. versionadded:: 0.27.0
        """
        if self.__update_obj is None:
            return False
        self.apply(obj=self.__update_obj, **kwargs)
        return True

    # endregion Update Methods

    def _get_mm100_obj_from_mm(self, value: UnitT | float, min_value: int = -9999) -> UnitMM100:
        """
        Gets a UnitMM100 object from mm.

        Args:
            value (UnitT | float): Units in mm or UnitT
            min_value (int, optional): The min value in ``1/100 mm`.
                If not != -9999 then value must be greater than or equal to ``min_val``.
                Defaults to -9999.

        Raises:
            ValueError: If min_val != -9999 and value < min_val

        Returns:
            UnitMM100: MM 100 units.
        """
        with contextlib.suppress(AttributeError):
            result = value.get_value_mm100()  # type: ignore
            if min_value != -9999 and result < min_value:
                raise ValueError("Value must be positive")
            return UnitMM100(result)
        result = UnitMM100.from_mm(value)  # type: ignore
        if min_value != -9999 and result.value < min_value:
            raise ValueError("Value must be positive")
        return result

    def _set_style_internal_events(self):
        self._fn_on_getting_cattribs = self._on_getting_cattribs
        self._fn_on_clearing = self._on_clearing
        self._fn_on_removing = self._on_removing
        self._fn_on_setting = self._on_setting
        self._fn_on_copying = self._on_copying
        self._fn_on_backing_up = self.on_property_backing_up
        self._fn_on_backed_up = self.on_property_backed_up
        self._fn_on_applying = self.on_applying
        self._fn_on_applied = self.on_applied

        self._events.on("internal_cattribs", self._fn_on_getting_cattribs)
        self._events.on(FormatNamedEvent.STYLE_CLEARING, self._fn_on_clearing)
        self._events.on(FormatNamedEvent.STYLE_REMOVING, self._fn_on_removing)
        self._events.on(FormatNamedEvent.STYLE_SETTING, self._fn_on_setting)
        self._events.on(FormatNamedEvent.STYLE_COPYING, self._fn_on_copying)
        self._events.on(FormatNamedEvent.STYLE_BACKING_UP, self._fn_on_backing_up)
        self._events.on(FormatNamedEvent.STYLE_BACKED_UP, self._fn_on_backed_up)
        self._events.on(FormatNamedEvent.STYLE_APPLYING, self._fn_on_applying)
        self._events.on(FormatNamedEvent.STYLE_APPLIED, self._fn_on_applied)

    # endregion Init

    # region Events

    def add_event_observer(self, *args: EventObserver) -> None:
        """
        Adds observers that gets their ``trigger`` method called when this class ``trigger`` method is called.

        Parameters:
            args (EventObserver): One or more observers to add.

        Returns:
            None:

        Note:
            Observers are removed automatically when they are out of scope.
        """
        self._events.add_observer(*args)

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

    def remove_event_observer(self, observer: EventObserver) -> bool:
        """
        Removes an observer.

        Args:
            observer (EventObserver): Observers to remove.

        Returns:
            bool: ``True`` if observer has been removed; Otherwise, ``False``.

        .. versionadded:: 0.30.1
        """
        return self._events.remove_observer(observer)

    # endregion Events

    # region style property methods

    def _get_properties(self) -> Dict[str, Any]:
        """Gets Key value pairs for the instance."""
        return self._dv

    def _get(self, key: str) -> Any:
        """Gets the property value"""
        return self._get_properties().get(key)

    def _set(self, key: str, val: Any) -> bool:
        """Sets a property value"""
        kvargs = KeyValCancelArgs("style_base", key=key, value=val)
        cargs = CancelEventArgs.from_args(kvargs)  # type: ignore
        self._events.trigger(FormatNamedEvent.STYLE_SETTING, kvargs)
        self._events.trigger(FormatNamedEvent.STYLE_MODIFYING, cargs)
        if kvargs.cancel:
            return False
        if cargs.cancel:
            return False
        data_values = self._get_properties()
        data_values[kvargs.key] = kvargs.value
        self._events.trigger(FormatNamedEvent.STYLE_SET, KeyValArgs.from_args(kvargs))  # type: ignore
        return True

    def _clear(self) -> None:
        """Clears all properties"""
        cargs = CancelEventArgs("style_base")
        self._events.trigger(FormatNamedEvent.STYLE_MODIFYING, cargs)
        self._events.trigger(FormatNamedEvent.STYLE_CLEARING, cargs)
        if cargs.cancel:
            return
        data_values = self._get_properties()
        data_values.clear()

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
            return False
        if self._has(key):
            data_values = self._get_properties()
            del data_values[key]
            return True
        return False

    def _del_attribs(self, *attribs: str) -> None:
        """Delete Attributes from instance if exist. Calls ``_on_deleting_attrib()``"""
        for attrib in attribs:
            if hasattr(self.__class__, attrib):
                kvargs = KeyValCancelArgs("style_base", key=attrib, value=getattr(self, attrib, None))
                self._on_deleting_attrib(self, kvargs)
                if kvargs.cancel:
                    continue
                delattr(self.__class__, attrib)  # type: ignore
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
        cargs.event_data = value
        self._events.trigger(FormatNamedEvent.STYLE_UPDATING, cargs)
        self._events.trigger(FormatNamedEvent.STYLE_MODIFYING, cargs)
        if cargs.cancel:
            return
        data_dict = self._get_properties()
        if isinstance(cargs.event_data, StyleBase):
            # pylint: disable=protected-access
            data_dict.update(cargs.event_data._dv)
            return
        data_dict.update(cargs.event_data)

    # endregion style property methods

    # region Services
    def support_service(self, *service: str) -> bool:
        """
        Gets if service is supported.

        Args:
            service: expandable list of service names of UNO services such as ``com.sun.star.text.TextFrame``.

        Returns:
            bool: ``True`` if service is supported; Otherwise, ``False``.
        """
        services = self._supported_services()
        return any(s in services for s in service)

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services for the style such as (``com.sun.star.style.ParagraphProperties``,)

        Raises:
            NotImplementedError: If not implemented in child class

        Returns:
            Tuple[str, ...]: Supported services
        """
        raise NotImplementedError

    def _is_valid_obj(self, obj: Any) -> bool:
        """
        Gets if ``obj`` supports one of the services required by style class

        Args:
            obj (Any): UNO object that must have requires service

        Returns:
            bool: ``True`` if has a required service; Otherwise, ``False``
        """
        return self._is_obj_service(obj)

    def _is_obj_service(self, obj: Any) -> bool:
        """
        Gets if ``obj`` supports one of the services required by style class

        Args:
            obj (Any): UNO object that must have requires service

        Returns:
            bool: ``True`` if has a required service; Otherwise, ``False``
        """
        if services := self._supported_services():
            return mInfo.Info.support_service(obj, *services)
        # if style class has no required services then return True
        return True

    def _print_not_valid_srv(self, method_name: str = ""):
        """
        Prints via ``Lo.print()`` notice that required service is missing

        Args:
            method_name (str, optional): Calling method name.
        """
        services = self._supported_services()
        rs_len = len(services)
        name = f".{method_name}()" if method_name else ""
        if rs_len == 0:
            mLo.Lo.print(f"{self.__class__.__name__}{name}: object is not valid.")
            return
        services = ", ".join(services)
        srv = "service" if rs_len == 1 else "services"
        mLo.Lo.print(f"{self.__class__.__name__}{name}: object must support {srv}: {services}")

    # endregion Services

    # region Internal Methods
    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        """Lo Safe Method."""
        # set properties. Can be overridden in child classes
        # may be useful to wrap in try statements in child classes
        try:
            mProps.Props.set(obj, **kwargs)
        except mEx.MultiError as multi_err:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in multi_err.errors:
                mLo.Lo.print(f"  {err}")

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
        return {} if cargs.cancel else cargs.event_data

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

    def apply(self, obj: Any, **kwargs: Any) -> None:
        """
        Applies styles to object

        Args:
            obj (Any): UNO Object that styles are to be applied.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Keyword Arguments:
            override_dv (Dic[str, Any], optional): if passed in this dictionary is used to set properties instead of internal dictionary of property values.
            validate (bool, optional): if ``False`` then ``obj`` is not validated. Defaults to ``True``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLIED` :eventref:`src-docs-event`

        Returns:
            None:

        Note:
            If Event data ``obj``, ``data_values`` or ``allow_update`` are changed then the new values are used.

            Add update object to instance if not already set and ``allow_update`` is ``True`` (default).

        .. versionchanged:: 0.27.0
            Event data is now a dictionary with keys ``source``, ``obj``, ``data_values`` and ``allow_update``.

        .. versionchanged:: 0.9.4
            Added ``validate`` keyword arguments.
        """
        allow_update = True
        validate = bool(kwargs.get("validate", True))
        if "override_dv" in kwargs:
            data_values = kwargs["override_dv"]
        else:
            data_values = self._get_properties()
        if len(data_values) > 0:
            if not validate or self._is_valid_obj(obj):
                cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
                event_data = {"source": self, "obj": obj, "data_values": data_values, "allow_update": allow_update}
                cargs.event_data = event_data
                self._events.trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
                if cargs.cancel:
                    return
                data_values = cargs.event_data["data_values"]
                the_obj = cargs.event_data["obj"]
                allow_update = cargs.event_data.get("allow_update", allow_update)
                with event_ctx(
                    (PropsNamedEvent.PROP_SETTING, _on_props_setting),
                    (PropsNamedEvent.PROP_SET, _on_props_set),
                    (PropsNamedEvent.PROP_SET_ERROR, _on_props_set_error),
                    source=self,
                    lo_observe=True,
                ):
                    self._props_set(the_obj, **data_values)
                # set for update.
                if allow_update is True and not self.has_update_obj():
                    self.set_update_obj(the_obj)
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
        return () if len(self._dv) == 0 else mProps.Props.make_props(**self._dv)

    # region Copy()
    @overload
    def copy(self: TStyleBase) -> TStyleBase: ...

    @overload
    def copy(self: TStyleBase, **kwargs) -> TStyleBase: ...

    def copy(self: TStyleBase, **kwargs) -> TStyleBase:
        """Gets a copy of instance as a new instance"""
        # sourcery skip: low-code-quality
        # pylint: disable=protected-access
        cargs = CancelEventArgs(self)
        self._events.trigger(FormatNamedEvent.STYLE_COPYING, cargs)
        if cargs.cancel:
            if cargs.handled:
                return cargs.event_data
            else:
                raise mEx.CancelEventError(cargs)
        new_class = self.__class__(**kwargs)
        new_class._prop_parent = self._prop_parent
        # pylint: disable=unused-variable
        if data_values := self._get_properties():
            # it is possible that that a new instance will have different property names then the current instance.
            # This can happen because this class inherits from MetaStyle.
            # if ne contains a _props attribute (tuple of prop names) then use them to remap keys.
            # For instance a key of BorderLength may become ParaBorderLength.

            key_map = None
            p_len = len(new_class._props)
            if p_len > 0 and p_len == len(self._props):
                key_map = {}
                for i, p_val in enumerate(self._props):
                    if p_val == "":
                        # some prop value may not be used in which case they are empty strings.
                        continue
                    if isinstance(p_val, PropPair):
                        nu_pair = cast(PropPair, new_class._props[i])
                        if p_val.first:
                            key_map[p_val.first] = nu_pair.first
                        if p_val.second:
                            key_map[p_val.second] = nu_pair.second
                    else:
                        key_map[p_val] = new_class._props[i]

            if key_map:
                for key, nu_val in key_map.items():
                    if self._has(key):
                        new_class._set(nu_val, self._get(key))
            else:
                new_class._update(self._get_properties())
        if self.__update_obj is not None:
            new_class.set_update_obj(self.__update_obj)
        return new_class

    # endregion Copy()

    # endregion Methods

    # region Backup/Restore

    def backup(self, obj: Any) -> None:
        """
        Backs up Attributes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (Any): Object to backup properties from.

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
            eargs = KeyValArgs.from_args(cargs)  # type: ignore
            self._events.trigger(FormatNamedEvent.STYLE_BACKED_UP, eargs)

    def restore(self, obj: Any, clear: bool = False) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (Any): Object to restore properties on.
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
            with event_ctx(
                (PropsNamedEvent.PROP_SETTING, _on_props_restore_setting),
                (PropsNamedEvent.PROP_SET, _on_props_restore_set),
                source=self,
                lo_observe=True,
            ):
                mProps.Props.set(obj, **self._dv_bak)

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
        # can be overridden in child classes.
        pass

    def on_property_set(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that is set.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overridden in child classes.
        pass

    def on_property_set_error(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that fails to set.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overridden in child classes.
        pass

    def on_property_backing_up(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be backup up during backup.

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args.
        """
        # can be overridden in child classes.
        pass

    def on_property_backed_up(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers before each property that is about to be set during backup.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overridden in child classes.
        pass

    def on_property_restore_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be set during restore

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args
        """
        # can be overridden in child classes.
        event_args.set("on_property_restore_setting", True)
        self.on_property_setting(source, event_args)

    def on_property_restore_set(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that has been set during restore

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overridden in child classes.
        event_args.set("on_property_restore_set", True)
        self.on_property_set(source, event_args)

    def on_applying(self, source: Any, event_args: CancelEventArgs) -> None:
        """
        Triggers before style/format is applied.

        Args:
            source (Any): Event Source.
            event_args (CancelEventArgs): Event Args
        """
        # can be overridden in child classes.
        pass

    def on_applied(self, source: Any, event_args: EventArgs) -> None:
        """
        Triggers after style/format is applied.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        # can be overridden in child classes.
        pass

    # endregion Event Methods

    # region Dunder Methods

    def __eq__(self, oth: object) -> bool:
        if not isinstance(oth, StyleBase):
            return NotImplemented
        result = False
        try:
            for key, value in self._get_properties().items():
                if oth._get(key) != value:
                    return False
            result = True
        except Exception:  # pylint: disable=broad-exception-caught
            return False
        return result

    # endregion Dunder Methods

    # region Named Container Methods

    def _container_get_service_name(self) -> str:
        raise NotImplementedError

    def _container_get_inst(self) -> XNameContainer:
        return mLo.Lo.create_instance_msf(
            XNameContainer,
            service_name=self._container_get_service_name(),
            msf=self._container_get_msf(),
            raise_err=True,
        )

    def _container_get_msf(self) -> XMultiServiceFactory | None:
        return None

    def _container_add_value(
        self,
        name: str,
        obj: Any,
        allow_update: bool = True,
        nc: XNameContainer | None = None,  # pylint: disable=invalid-name
    ) -> None:
        if nc is None:
            nc = self._container_get_inst()
        if nc.hasByName(name):
            if allow_update:
                nc.replaceByName(name, obj)
                return
        else:
            nc.insertByName(name, obj)

    def _container_get_unique_el_name(
        self, prefix: str, nc: XNameContainer | None = None  # pylint: disable=invalid-name
    ) -> str:
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

    def _container_get_value(self, name: str, nc: XNameContainer | None = None) -> Any:  # pylint: disable=invalid-name
        if not name:
            raise ValueError("Name is empty value. Expected a string name.")
        if nc is None:
            nc = self._container_get_inst()
        return nc.getByName(name) if nc.hasByName(name) else None

    # endregion Named Container Methods

    # region Properties

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instance has any attributes set."""
        return len(self._dv) > 0

    @property
    def prop_has_backup(self) -> bool:
        """Gets If instance has backup data set."""
        return False if self._dv_bak is None else len(self._dv_bak) > 0

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
        # pylint: disable=no-member
        return self._internal_events  # type: ignore

    # endregion Properties


# endregion Style Base Class


# region Module Internal Helper Classes
class _StyleMultiArgs:
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
    style: StyleT
    args: _StyleMultiArgs | None


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
    def _get_service_name(self) -> str:
        raise NotImplementedError

    def _supported_services(self) -> Tuple[str, ...]:
        raise NotImplementedError

    def _set_style_internal_events(self) -> None:
        super()._set_style_internal_events()

        self._fn_on_multi_style_setting = self._on_multi_style_setting
        self._fn_on_multi_style_set = self._on_multi_style_set
        self._fn_on_multi_style_removing = self._on_multi_style_removing
        self._fn_on_multi_style_removed = self._on_multi_style_removed
        self._fn_on_multi_style_updating = self._on_multi_style_updating
        self._fn_on_multi_style_updated = self._on_multi_style_updated
        self._fn_on_multi_child_style_applying = self._on_multi_child_style_applying
        self._fn_on_multi_child_style_applied = self._on_multi_child_style_applied

        self._events.on(FormatNamedEvent.MULTI_STYLE_SETTING, self._fn_on_multi_style_setting)
        self._events.on(FormatNamedEvent.MULTI_STYLE_SET, self._fn_on_multi_style_set)
        self._events.on(FormatNamedEvent.MULTI_STYLE_REMOVING, self._fn_on_multi_style_removing)
        self._events.on(FormatNamedEvent.MULTI_STYLE_REMOVED, self._fn_on_multi_style_removed)
        self._events.on(FormatNamedEvent.MULTI_STYLE_UPDATING, self._fn_on_multi_style_updating)
        self._events.on(FormatNamedEvent.MULTI_STYLE_UPDATED, self._fn_on_multi_style_updated)
        self._events.on(FormatNamedEvent.STYLE_MULTI_CHILD_APPLYING, self._fn_on_multi_child_style_applying)
        self._events.on(FormatNamedEvent.STYLE_MULTI_CHILD_APPLIED, self._fn_on_multi_child_style_applied)

    # region Update Methods

    def set_update_obj(self, obj: Any) -> None:
        """
        Sets the update object for the styles instances.

        Args:
            obj (Any): Object used to apply style to when update is called.
        """
        super().set_update_obj(obj)
        styles = self._get_multi_styles()
        for style, _ in styles.values():
            style.set_update_obj(obj)

    def update(self, **kwargs: Any) -> bool:
        """
        Applies the styles to the update object.

        Returns:
            bool: Returns ``True`` if the update object is set and the style is applied; Otherwise, ``False``.
            kwargs: Expandable list of key value pairs that may be used in child classes.
        """
        result = True
        styles = self._get_multi_styles()
        for style, _ in styles.values():
            result = result and style.update(**kwargs)
        result = result and super().update(**kwargs)
        return result

    # endregion Update Methods

    # region apply()

    def _apply_direct(self, obj: Any, **kwargs) -> None:
        """Calls super apply directly"""
        super().apply(obj, **kwargs)

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies style of current instance and all other internal style instances.

        Args:
            obj (Any): UNO Object that styles are to be applied.
        """
        styles = self._get_multi_styles()
        for key, info in styles.items():
            kvargs = KeyValCancelArgs("StyleMulti.apply", key=key, value=info)
            self._events.trigger(FormatNamedEvent.STYLE_MULTI_CHILD_APPLYING, kvargs)
            if kvargs.cancel:
                continue
            style, key_word = info
            if key_word:
                style.apply(obj, **key_word.kwargs)
            else:
                style.apply(obj)
            self._events.trigger(FormatNamedEvent.STYLE_MULTI_CHILD_APPLIED, KeyValArgs.from_args(kvargs))  # type: ignore
        # apply this instance properties after all others styles.
        # allows this instance to overwrite properties set by multi styles if needed.
        super().apply(obj, **kwargs)

    # endregion apply()

    # region Copy()
    @overload
    def copy(self: TStyleMulti) -> TStyleMulti: ...

    @overload
    def copy(self: TStyleMulti, **kwargs) -> TStyleMulti: ...

    def copy(self: TStyleMulti, **kwargs) -> TStyleMulti:
        """Gets a copy of instance as a new instance"""
        # pylint: disable=protected-access
        instance_copy = super().copy(**kwargs)
        styles = self._get_multi_styles()
        for key, info in styles.items():
            style, key_words = info
            if key_words:
                instance_copy._set_style(key, style.copy(**kwargs), **key_words.kwargs)
            else:
                instance_copy._set_style(key, style.copy(**kwargs))
        return instance_copy

    # endregion Copy()

    def __eq__(self, oth: object) -> bool:
        if not isinstance(oth, StyleMulti):
            return NotImplemented
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
            if not result:
                break
        return result

    # endregion Overrides

    # region Internal Methods

    def _set_style(self, key: str, style: StyleT, *attrs, **kwargs) -> None:
        """
        Sets style

        Args:
            key (str): key store style info
            style (StyleBase): style
            attrs: Expandable list attributes that style sets.
                The values added here are added when get_attrs() method is called.
                This is used for backup and restore in Write Module.
            kwargs: Expandable key value args to that are to be passed to style when ``apply_style()`` is called.
        """
        # pylint: disable=protected-access
        kvargs = KeyValCancelArgs("_set_style", key=key, value=style)
        self._events.trigger(FormatNamedEvent.MULTI_STYLE_SETTING, kvargs)
        if kvargs.cancel:
            return
        styles = self._get_multi_styles()
        if style._prop_parent is None:  # type: ignore
            style._prop_parent = self  # type: ignore
        if len(attrs) + len(kwargs) == 0:
            styles[key] = _StyleInfo(style, None)
        else:
            styles[key] = _StyleInfo(style, _StyleMultiArgs(*attrs, **kwargs))
        self._events.trigger(FormatNamedEvent.MULTI_STYLE_SET, KeyValArgs.from_args(kvargs))  # type: ignore

    def _update_style(self, value: StyleMulti) -> None:
        # pylint: disable=protected-access
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
            return False
        result = False
        styles = self._get_multi_styles()
        if key in styles:
            del styles[key]
            result = True
            # only trigger if removed
            self._events.trigger(FormatNamedEvent.MULTI_STYLE_REMOVED, EventArgs.from_args(cargs))
        return result

    def _get_style(self, key: str) -> _StyleInfo | None:
        return self._get_multi_styles().get(key)

    def _get_style_inst(self, key: str) -> StyleT | None:
        style = self._get_style(key)
        return None if style is None else style.style

    def _has_style(self, key: str) -> bool:
        return key in self._get_multi_styles()

    def _get_multi_styles(self) -> Dict[str, _StyleInfo]:
        return self._styles

    # endregion Internal Methods

    # region Methods
    def backup(self, obj: Any) -> None:
        """
        Backs up Attributes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (Any): Object to backup properties from.

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

    def restore(self, obj: Any, clear: bool = False) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (Any): Object to restore properties on.
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
            if styles := self._get_multi_styles():
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
        return False if self._dv_bak is None else len(self._dv_bak) > 0

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

    def _is_valid_obj(self, obj: Any) -> bool:
        if mLo.Lo.is_uno_interfaces(obj, "com.sun.star.style.XStyle"):
            return self._is_obj_service(obj)
        # pylint: disable=unused-variable
        if self._is_obj_service(obj):
            return True
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as multi_err:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in multi_err.errors:
                mLo.Lo.print(f"  {err}")

    # region apply()
    # pylint: disable=arguments-differ
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (Any): UNO Object such as a Document, Spreadsheet, or ``XStyle``

        Returns:
            None:
        """

        if mLo.Lo.is_uno_interfaces(obj, "com.sun.star.style.XStyle"):
            if not self._is_obj_service(obj):
                self._print_not_valid_srv(method_name="apply")
                return
            props = mLo.Lo.qi(XPropertySet, obj)
            if props is None:
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
            props = self.get_style_props(obj)
        super().apply(props, **kwargs)

    # endregion apply()

    # region copy()
    @overload
    def copy(self: _TStyleModifyMulti) -> _TStyleModifyMulti: ...

    @overload
    def copy(self: _TStyleModifyMulti, **kwargs) -> _TStyleModifyMulti: ...

    def copy(self: _TStyleModifyMulti, **kwargs) -> _TStyleModifyMulti:
        """Gets a copy of instance as a new instance"""
        inst_copy = super().copy(**kwargs)
        inst_copy.prop_style_name = self.prop_style_name
        return inst_copy

    # endregion copy()

    # endregion Overrides

    # region internal methods

    def _is_valid_doc(self, obj: Any) -> bool:  # pylint: disable=unused-argument
        return True
        # return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)

    def _get_style_family_name(self) -> str:
        if not self._style_family_name:
            raise ValueError("Must set internal property _style_family_name")
        return self._style_family_name

    # endregion internal methods

    # region Methods

    def get_style_props(self, doc: Any) -> XPropertySet:
        """
        Gets the Style Properties

        Args:
            doc (Any): UNO Document Object.

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
        # sourcery skip: dict-assign-update-to-union
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
        doc = mLo.Lo.this_component
        if doc is None:
            # most likely in headless mode with option dynamic set to True
            doc = mLo.Lo.lo_component
        return mInfo.Info.get_style_props(
            doc=doc, family_style_name=self._get_family_style_name(), prop_set_nm=self.prop_name
        )

    # endregion methods

    # region internal methods
    def _get_family_style_name(self) -> str:
        raise NotImplementedError

    def _get_property_name(self) -> str:
        # sourcery skip: raise-from-previous-error
        try:
            return self._style_property_name  # type: ignore
        except AttributeError:
            # pylint: disable=raise-missing-from
            raise NotImplementedError

    # endregion internal methods

    # region Overrides
    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event_args)

    # endregion Overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[TStyleName], obj: Any) -> TStyleName: ...

    @overload
    @classmethod
    def from_obj(cls: Type[TStyleName], obj: Any, **kwargs) -> TStyleName: ...

    @classmethod
    def from_obj(cls: Type[TStyleName], obj: Any, **kwargs) -> TStyleName:
        """
        Gets instance from object

        Args:
            obj (Any): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TStyleName: ``TStyleName`` instance that represents ``obj`` style.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)

        if mInfo.Info.support_service(obj, "com.sun.star.style.CellStyle"):
            cell_style = cast("CellStyle", obj)
            name = cell_style.getName()
            inst.prop_name = name
            return inst
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        if name := mProps.Props.get(obj, inst._get_property_name(), ""):
            inst.prop_name = name
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
def _on_props_setting(
    source: Any, event_args: KeyValCancelArgs, *args, **kwargs  # pylint: disable=unused-argument
) -> None:
    # pylint: disable=protected-access
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_setting(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_APPLYING, event_args)


def _on_props_set(source: Any, event_args: KeyValArgs, *args, **kwargs) -> None:  # pylint: disable=unused-argument
    # pylint: disable=protected-access
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_set(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_APPLIED, event_args)


def _on_props_set_error(
    source: Any, event_args: KeyValCancelArgs, *args, **kwargs  # pylint: disable=unused-argument
) -> None:
    # pylint: disable=protected-access
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_set_error(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_ERROR, event_args)


def _on_props_restore_setting(
    source: Any, event_args: KeyValCancelArgs, *args, **kwargs  # pylint: disable=unused-argument
) -> None:
    # pylint: disable=protected-access
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_restore_setting(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_RESTORING, event_args)


def _on_props_restore_set(
    source: Any, event_args: KeyValArgs, *args, **kwargs  # pylint: disable=unused-argument
) -> None:
    # pylint: disable=protected-access
    instance = cast(StyleBase, event_args.event_source)
    instance.on_property_restore_set(source, event_args)
    instance._events.trigger(FormatNamedEvent.STYLE_PROPERTY_RESTORED, event_args)


# endregion Props Property event handlers

__all__ = ("StyleBase", "StyleMulti", "StyleModifyMulti", "StyleName", "TStyleBase", "TStyleMulti", "TStyleName")
