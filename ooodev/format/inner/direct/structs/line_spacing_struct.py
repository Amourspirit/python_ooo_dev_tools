"""
Module for Shadow format (``LineSpacing``) struct.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations

from typing import Any, Dict, Tuple, Type, TypeVar, cast, overload, TYPE_CHECKING
from enum import Enum

from ooo.dyn.style.line_spacing import LineSpacing as UnoLineSpacing

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.format_named_event import FormatNamedEvent
from ooodev.format.inner.direct.structs.struct_base import StructBase
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.units.unit_convert import UnitConvert, UnitLength
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
# endregion Import

_TLineSpacingStruct = TypeVar(name="_TLineSpacingStruct", bound="LineSpacingStruct")


class ModeKind(Enum):
    """Mode Kind for line spacing"""

    # Enum value, mode, default
    SINGLE = (0, 0, 100)  # zero value,
    """Single Line Spacing ``1mm``"""
    LINE_1_15 = (1, 0, 115)  # zero value, value
    """Line Spacing ``1.15mm``"""
    LINE_1_5 = (2, 0, 150)
    """Line Spacing ``1.5mm``"""
    DOUBLE = (3, 0, 200)
    """Double line spacing ``2mm``"""
    PROPORTIONAL = (4, 0, 0)  # PERCENTAGE, No conversion onf height value 98 % = 98 MM100
    """Proportional line spacing"""
    AT_LEAST = (5, 1, 0)  # IN 1/100 MM
    """At least line spacing"""
    LEADING = (6, 2, 0)  # in 1/100 MM
    """Leading Line Spacing"""
    FIXED = (7, 3, 0)  # in 1/100 MM
    """Fixed Line Spacing"""

    def __int__(self) -> int:
        return self.value[2]

    def get_mode(self) -> int:
        return self.value[1]

    def get_enum_val(self) -> int:
        return self.value[0]

    @staticmethod
    def from_uno(ls: UnoLineSpacing) -> ModeKind:
        """Converts UNO ``LineSpacing`` struct to ``ModeKind`` enum."""
        mode = ls.Mode
        val = ls.Height
        if mode == 0:
            if val == 100:
                return ModeKind.SINGLE
            if val == 115:
                return ModeKind.LINE_1_15
            if val == 150:
                return ModeKind.LINE_1_5
            return ModeKind.DOUBLE if val == 200 else ModeKind.PROPORTIONAL
        if mode == 1:
            return ModeKind.AT_LEAST
        if mode == 2:
            return ModeKind.LEADING
        if mode == 3:
            return ModeKind.FIXED
        raise ValueError("Unable to convert uno LineSpacing object to ModeKind Enum")


class LineSpacingStruct(StructBase):
    """
    Line Spacing struct
    """

    # region init

    def __init__(self, mode: ModeKind = ModeKind.SINGLE, value: int | float | UnitT = 0) -> None:
        """
        Constructor

        Args:
            mode (LineMode, optional): This value specifies the way the spacing is specified.
            value (Real, UnitT, optional): This value specifies the spacing in regard to Mode.

        Raises:
            ValueError: If ``value`` are less than zero.

        Note:
            If ``LineMode`` is ``SINGLE``, ``LINE_1_15``, ``LINE_1_5``, or ``DOUBLE`` then ``value`` is ignored.

            If ``LineMode`` is ``AT_LEAST``, ``LEADING``, or ``FIXED`` then ``value`` is a float (``in mm units``)
            or :ref:`proto_unit_obj`

            If ``LineMode`` is ``PROPORTIONAL`` then value is an int representing percentage.
            For example ``95`` equals ``95%``, ``130`` equals ``130%``
        """

        self._line_mode = mode
        self._mode = mode.get_mode()
        self._value = int(mode)
        enum_val = mode.get_enum_val()
        if mode == ModeKind.PROPORTIONAL:
            # no conversion
            try:
                # just in case passed in as a UnitT
                self._value = round(value.value)  # type: ignore
            except AttributeError:
                self._value = int(value)  # type: ignore

        elif enum_val >= 5:
            try:
                self._value = cast(int, value.get_value_mm100())  # type: ignore
            except AttributeError:
                self._value = round(UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.MM100))  # type: ignore

        if self._value < 0:
            raise ValueError("mode must be a positive number")

        super().__init__()

    # endregion init

    # region methods
    def __eq__(self, other: object) -> bool:
        ls2 = None
        if isinstance(other, LineSpacingStruct):
            ls2 = other.get_uno_struct()
        elif getattr(other, "typeName", None) == "com.sun.star.style.LineSpacing":
            ls2 = other
        if ls2:
            ls1 = self.get_uno_struct()
            return ls1.Height == ls2.Height and ls1.Mode == ls2.Mode  # type: ignore
        return False

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "ParaLineSpacing"
        return self._property_name

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attributes
        """
        return (self._get_property_name(),)

    # region copy()
    @overload
    def copy(self: _TLineSpacingStruct) -> _TLineSpacingStruct: ...

    @overload
    def copy(self: _TLineSpacingStruct, **kwargs) -> _TLineSpacingStruct: ...

    def copy(self: _TLineSpacingStruct, **kwargs) -> _TLineSpacingStruct:
        nu = self.__class__(mode=self._mode, height=self._value, **kwargs)  # type: ignore
        if dv := self._get_properties():
            nu._update(dv)
        return nu

    # endregion copy()

    # region apply()

    @overload
    def apply(self, obj: Any, *, keys: Dict[str, str]) -> None: ...

    @overload
    def apply(self, obj: Any) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object that contains a ``LineSpacing`` property.
            keys: (Dict[str, str], optional): key map for properties.
                Can be ``spacing`` which maps to ``ParaLineSpacing`` by default.

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
            self._print_not_valid_srv("apply()")
            return
        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        if cargs.cancel:
            return
        self._events.trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        keys = {"spacing": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        key = keys["spacing"]
        mProps.Props.set(obj, **{key: self.get_uno_struct()})
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

    # endregion apply()

    def get_uno_struct(self) -> UnoLineSpacing:
        """
        Gets UNO ``Gradient`` from instance.

        Returns:
            Gradient: ``Gradient`` instance
        """
        return UnoLineSpacing(Mode=self._mode, Height=self._value)

    # region from_line_spacing()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TLineSpacingStruct], ln_spacing: UnoLineSpacing) -> _TLineSpacingStruct: ...

    @overload
    @classmethod
    def from_uno_struct(
        cls: Type[_TLineSpacingStruct], ln_spacing: UnoLineSpacing, **kwargs
    ) -> _TLineSpacingStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TLineSpacingStruct], ln_spacing: UnoLineSpacing, **kwargs) -> _TLineSpacingStruct:
        """
        Converts a UNO ``LineSpacing`` struct into a ``LineSpacingStruct``

        Args:
            ln_spacing (UnoLineSpacing): UNO ``LineSpacing`` object.

        Returns:
            LineSpacingStruct: ``LineSpacingStruct`` set with Line spacing properties.
        """
        inst = cls(**kwargs)
        inst._mode = ln_spacing.Mode
        inst._value = ln_spacing.Height
        inst._line_mode = ModeKind.from_uno(ln_spacing)
        return inst

    # endregion from_line_spacing()

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STRUCT
        return self._format_kind_prop

    @property
    def prop_mode(self) -> ModeKind:
        """Gets mode value"""
        return self._line_mode

    @property
    def prop_value(self) -> int:
        """Gets the spacing value in regard to Mode"""
        return self._value

    @property
    def default(self: _TLineSpacingStruct) -> _TLineSpacingStruct:  # type: ignore[misc]
        """Gets empty Line Spacing. Static Property."""
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(ModeKind.SINGLE, 0, _cattribs=self._get_internal_cattribs())  # type: ignore
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion Properties
