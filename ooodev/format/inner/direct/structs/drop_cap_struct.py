"""
Module for ``DropCapFormat`` struct.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Dict, Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING

from ooo.dyn.style.drop_cap_format import DropCapFormat

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.format_named_event import FormatNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.struct_base import StructBase
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.utils import props as mProps
from ooodev.utils.data_type.byte import Byte

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TDropCapStruct = TypeVar(name="_TDropCapStruct", bound="DropCapStruct")


class DropCapStruct(StructBase):
    """
    Paragraph Drop Cap

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, *, count: int = 0, distance: int | UnitT = 0, lines: int = 0) -> None:
        """
        Constructor

        Args:
            count (int, optional): Specifies the number of characters in the drop cap. Must be from ``0`` to ``255``. Defaults to ``0``
            distance (int, UnitT, optional): Specifies the distance between the drop cap in the following text in ``1/100th mm`` or :ref:`proto_unit_obj`. Defaults to ``0``
            lines (int, optional): Specifies the number of lines used for a drop cap. Must be from ``0`` to ``255``. Defaults to ``0``

        Returns:
            None:

        Note:
            If argument ``type`` is ``None`` then all other argument are ignored
        """
        b_count = Byte(count)
        count = b_count.value
        b_lines = Byte(lines)
        lines = b_lines.value
        try:
            dist = distance.get_value_mm100()  # type: ignore
        except AttributeError:
            dist = int(distance)  # type: ignore
        if dist < 0:
            raise ValueError("distance arg must be a positive value.")

        init_vals = {"Count": count, "Distance": dist, "Lines": lines}

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

    def _is_valid_obj(self, obj: Any) -> bool:
        return mProps.Props.has(obj, self._get_property_name())

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, keys: Dict[str, str]) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:  # type: ignore
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
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLIED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        # sourcery skip: dict-assign-update-to-union
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_srv("apply")
            return

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        if cargs.cancel:
            return
        self._events.trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        keys = {"prop": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        key = keys["prop"]
        dcf = self.get_uno_struct()
        mProps.Props.set(obj, **{key: dcf})
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

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
    def from_obj(cls: Type[_TDropCapStruct], obj: Any) -> _TDropCapStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TDropCapStruct], obj: Any, **kwargs) -> _TDropCapStruct: ...

    @classmethod
    def from_obj(cls: Type[_TDropCapStruct], obj: Any, **kwargs) -> _TDropCapStruct:
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
        except mEx.PropertyNotFoundError as e:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property") from e

        return cls.from_uno_struct(dcf, **kwargs)

    # endregion from_obj()

    # region from_drop_cap_format()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TDropCapStruct], dcf: DropCapFormat) -> _TDropCapStruct: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TDropCapStruct], dcf: DropCapFormat, **kwargs) -> _TDropCapStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TDropCapStruct], dcf: DropCapFormat, **kwargs) -> _TDropCapStruct:
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

    def fmt_distance(self: _TDropCapStruct, value: int | UnitT) -> _TDropCapStruct:
        """
        Gets a copy of instance with distance set.

        Args:
            value (int, UnitT): Distance value in ``1/100th mm`` or :ref:`proto_unit_obj`.

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
    def prop_distance(self) -> UnitMM100:
        """Gets/Sets the distance between the drop cap in the following text."""
        return UnitMM100(self._get("Distance"))

    @prop_distance.setter
    def prop_distance(self, value: int | UnitT) -> None:
        try:
            val = cast(int, value.get_value_mm100())  # type: ignore
        except AttributeError:
            val = cast(int, value)
        if val < 0:
            raise ValueError("value must be a positive number.")
        self._set("Distance", val)

    # endregion properties
