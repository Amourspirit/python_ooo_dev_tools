"""
Module for ``DropCapFormat`` struct.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Dict, Tuple, Type, cast, overload, TypeVar

import uno
from ....events.event_singleton import _Events
from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....utils.data_type.byte import Byte
from ....utils.type_var import T
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent


from ooo.dyn.style.drop_cap_format import DropCapFormat

_TDropCapStruct = TypeVar(name="_TDropCapStruct", bound="DropCapStruct")


class DropCapStruct(StyleBase):
    """
    Paragraph Drop Cap

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, *, count: int = 0, distance: int = 0, lines: int = 0) -> None:
        """
        Constructor

        Args:
            count (int, optional): Specifies the number of characters in the drop cap. Must be from ``0`` to ``255``. Defaults to ``0``
            distance (int, optional): Specifies the distance between the drop cap in the following text. Defaults to ``0``
            lines (int, optional): Specifies the number of lines used for a drop cap. Must be from ``0`` to ``255``. Defaults to ``0``

        Returns:
            None:

        Note:
            If argument ``type`` is ``None`` then all other argument are ignored
        """
        bcount = Byte(count)
        count = bcount.value
        blines = Byte(lines)
        lines = blines.value
        if distance < 0:
            raise ValueError("distance arg must be a positive number.")

        init_vals = {"Count": count, "Distance": distance, "Lines": lines}

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "DropCapFormat"
        return self._property_name

    def _is_valid_obj(self, obj: object) -> bool:
        return mProps.Props.has(obj, self._get_property_name())

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    @overload
    def apply(self, obj: object, keys: Dict[str, str]) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        If a DropCap instance with the same position is existing it is updated;
        Otherwise, a new DropCap is added.

        Args:
            obj (object): UNO object.
            keys (Dict[str, str], optional): Property key, value items that map properties.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_obj("apply")
            return

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        keys = {"prop": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        key = keys["prop"]
        dcf = self.get_uno_struct()
        mProps.Props.set(obj, **{key: dcf})
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

        # mProps.Props.set(obj, **{key: tuple(tss_lst)})

    # endregion apply()

    def get_uno_struct(self) -> DropCapFormat:
        """
        Gets ``DropCapFormat`` from instance

        Returns:
            DropCapFormat: ``DropCapFormat`` instance
        """
        return DropCapFormat(Lines=self._get("Lines"), Count=self._get("Count"), Distance=self._get("Distance"))

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TDropCapStruct], obj: object) -> _TDropCapStruct:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TDropCapStruct], obj: object, **kwargs) -> _TDropCapStruct:
        ...

    @classmethod
    def from_obj(cls: Type[_TDropCapStruct], obj: object, **kwargs) -> _TDropCapStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            DropCap: ``DropCap`` instance that represents ``obj`` Drop cap format properties.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            dcf = cast(DropCapFormat, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_drop_cap_format(dcf, **kwargs)

    # endregion from_obj()

    # region from_drop_cap_format()
    @overload
    @classmethod
    def from_drop_cap_format(cls: Type[_TDropCapStruct], dcf: DropCapFormat) -> _TDropCapStruct:
        ...

    @overload
    @classmethod
    def from_drop_cap_format(cls: Type[_TDropCapStruct], dcf: DropCapFormat, **kwargs) -> _TDropCapStruct:
        ...

    @classmethod
    def from_drop_cap_format(cls: Type[_TDropCapStruct], dcf: DropCapFormat, **kwargs) -> _TDropCapStruct:
        """
        Converts a ``DropCapFormat`` Stop instance to a ``DropCap``

        Args:
            dcf (DropCapFormat): UNO Drop Cap Format

        Returns:
            DropCap: ``DropCap`` set with Drop Cap Format properties
        """
        return cls(count=dcf.Count, distance=dcf.Distance, lines=dcf.Lines, **kwargs)

    # endregion from_drop_cap_format()

    # endregion methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, DropCapStruct):
            return (
                self._get("Count") == oth._get("Count")
                and self._get("Distance") == oth._get("Distance")
                and self._get("Lines") == oth._get("Lines")
            )
        if hasattr(oth, "typeName") and getattr(oth, "typeName") == "com.sun.star.style.DropCapFormat":
            dcf = cast(DropCapFormat, oth)
            return (
                self._get("Count") == dcf.Count
                and self._get("Distance") == dcf.Distance
                and self._get("Lines") == dcf.Lines
            )
        return NotImplemented

    # endregion dunder methods

    # region format methods
    def fmt_count(self: _TDropCapStruct, value: int) -> _TDropCapStruct:
        """
        Gets a copy of instance with count set.

        Args:
            value (int): Count value.

        Returns:
            DropCap: ``DropCap`` instance
        """
        cp = self.copy()
        cp.prop_count = value
        return cp

    def fmt_distance(self: _TDropCapStruct, value: int) -> _TDropCapStruct:
        """
        Gets a copy of instance with distance set.

        Args:
            value (int): Distance value.

        Returns:
            DropCap: ``DropCap`` instance
        """
        cp = self.copy()
        cp.prop_distance = value
        return cp

    def fmt_lines(self: _TDropCapStruct, value: int) -> _TDropCapStruct:
        """
        Gets a copy of instance with lines set.

        Args:
            value (int): Lines value.

        Returns:
            DropCap: ``DropCap`` instance
        """
        cp = self.copy()
        cp.prop_lines = value
        return cp

    # endregion format methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    @property
    def prop_count(self) -> int:
        """Gets/Sets the number of characters in the drop cap. Must be from 0 to 255."""
        return self._get("Count")

    @prop_count.setter
    def prop_count(self, value: int) -> None:
        val = Byte(value)
        self._set("Count", val.value)

    @property
    def prop_lines(self) -> int:
        """Gets/Sets the number of lines used for a drop cap. Must be from 0 to 255."""
        return self._get("Lines")

    @prop_lines.setter
    def prop_lines(self, value: int) -> None:
        val = Byte(value)
        self._set("Lines", val.value)

    @property
    def prop_distance(self) -> int:
        """Gets/Sets the distance between the drop cap in the following text."""
        return self._get("Distance")

    @prop_distance.setter
    def prop_distance(self, value: int) -> None:
        if value < 0:
            raise ValueError("value must be a positive number.")
        self._set("Distance", value)

    # endregion properties
