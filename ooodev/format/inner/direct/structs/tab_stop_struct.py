# region Import
from __future__ import annotations
from typing import Any, Dict, Iterable, Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING
from enum import Enum

import uno
from com.sun.star.beans import XPropertySet

from ooo.dyn.style.tab_align import TabAlign
from ooo.dyn.style.tab_stop import TabStop

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.format_named_event import FormatNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.struct_base import StructBase
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.loader import lo as mLo
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import

_TTabStopStruct = TypeVar("_TTabStopStruct", bound="TabStopStruct")


class FillCharKind(Enum):
    """Tab Fill Char"""

    NONE = " "
    DECIMAL = "."
    DASH = "-"
    UNDER_SCORE = "_"

    def __str__(self) -> str:
        return self.value


class TabStopStruct(StructBase):
    """
    Paragraph Tab

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        position: float | UnitT = 0.0,
        align: TabAlign = TabAlign.LEFT,
        decimal_char: str = ".",
        fill_char: FillCharKind | str = FillCharKind.NONE,
    ) -> None:
        """
        Constructor

        Args:
            position (float, UnitT, optional): Specifies the position of the tabulator in relation to the left border (in ``mm`` units) or :ref:`proto_unit_obj`.
                Defaults to ``0.0``
            align (TabAlign, optional): Specifies the alignment of the text range before the tabulator. Defaults to ``TabAlign.LEFT``
            decimal_char (str, optional): Specifies which delimiter is used for the decimal.
                Argument is expected to be a single character string.
                This argument is only used when ``align`` is set to ``TabAlign.DECIMAL``.
            fill_char (FillCharKind, str, optional): specifies the character that is used to fill up the space between the text in the text range and the tabulators.
                If string value then argument is expected to be a single character string.
                Defaults to ``FillCharKind.NONE``

        Returns:
            None:

        Note:
            If argument ``type`` is ``None`` then all other argument are ignored
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        # in Writer when a tab is added it is static. That is it remains when subsequent paragraphs are added.
        # existing paragraph can can add or remove Tabs. This will not affect the tabs of other existing paragraphs.
        # most likely it would be best practice not to back up and restore this property
        # (ParaLineSpacing) when appending paragraphs in Writer.

        init_vals = {"FillChar": str(fill_char)[:1], "Alignment": align, "DecimalChar": decimal_char[:1]}
        try:
            init_vals["Position"] = position.get_value_mm100()  # type: ignore
        except AttributeError:
            init_vals["Position"] = UnitConvert.convert_mm_mm100(position)  # type: ignore
        if align != TabAlign.DECIMAL:
            init_vals["DecimalChar"] = " "

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
            self._property_name = "ParaTabStops"
        return self._property_name

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

        If a Tab instance with the same position is existing it is updated;
        Otherwise, a new Tab is added.

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
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
            self._print_not_valid_srv("apply()")
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

        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, key))
        match = -1
        # often Writer will change Position values, 800 to 801 etc.
        # for this reason using a range to check plus or minus 2
        pos = int(self._get("Position"))
        pos_rng = range(pos - 2, pos + 3)  # plus or minus 2
        for i, ts in enumerate(tss):
            if pos == ts.Position:
                # also covers 0
                match = i
                break
            if ts.Position in pos_rng:
                match = i
                break

        ts = self.get_uno_struct()
        tss_lst = list(tss)
        if match >= 0:
            tss_lst[match] = ts
        else:
            tss_lst.append(ts)

        self._set_obj_tabs(obj, tss_lst, key)
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

        # mProps.Props.set(obj, **{key: tuple(tss_lst)})

    # endregion apply()

    def get_uno_struct(self) -> TabStop:
        """
        Gets UNO ``TabStop`` from instance.

        Returns:
            TabStop: ``TabStop`` instance
        """
        dec = self._get("DecimalChar") if self.prop_align == TabAlign.DECIMAL else " "
        return TabStop(
            Position=self._get("Position"),
            Alignment=self._get("Alignment"),
            DecimalChar=dec,
            FillChar=self._get("FillChar"),
        )

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTabStopStruct], obj: Any) -> _TTabStopStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTabStopStruct], obj: Any, **kwargs) -> _TTabStopStruct: ...

    @classmethod
    def from_obj(cls: Type[_TTabStopStruct], obj: Any, index: int = 0, **kwargs) -> _TTabStopStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object
            index (int, optional): Index of tab stop. Default ``0``.

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            Tab: ``Tab`` instance that represents ``obj`` Tab properties.
        """
        # sourcery skip: raise-from-previous-error
        # this nu is only used to get Property Name
        # pylint: disable=protected-access
        # pylint: disable=raise-missing-from
        # pylint: disable=unsubscriptable-object
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()
        try:
            tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")
        # can expect for ts to contain at least one TabAlign
        ts = tss[index]

        return cls.from_uno_struct(ts, **kwargs)

    # endregion from_obj()

    # region from_tab_stop()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TTabStopStruct], ts: TabStop) -> _TTabStopStruct: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TTabStopStruct], ts: TabStop, **kwargs) -> _TTabStopStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TTabStopStruct], ts: TabStop, **kwargs) -> _TTabStopStruct:
        """
        Converts a Tab Stop instance to a Tab

        Args:
            ts (TabStop): Tab stop

        Returns:
            Tab: Tab set with Tab Stop properties
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        inst._set("FillChar", ts.FillChar)
        inst._set("Alignment", ts.Alignment)
        inst._set("DecimalChar", ts.DecimalChar)
        inst._set("Position", ts.Position)
        return inst

    # endregion from_tab_stop()

    def _set_obj_tabs(self, obj: Any, tabs: Iterable[TabStop], prop: str = "") -> None:
        if not prop:
            prop = self._get_property_name()
        prop_set = mLo.Lo.qi(XPropertySet, obj, raise_err=True)
        seq = uno.Any("[]com.sun.star.style.TabStop", tabs)  # type: ignore
        uno.invoke(prop_set, "setPropertyValue", (prop, seq))  # type: ignore

    # endregion methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, TabStopStruct):
            return (
                self._get("Position") == oth._get("Position")
                and self._get("Alignment") == oth._get("Alignment")
                and self._get("DecimalChar") == oth._get("DecimalChar")
                and self._get("FillChar") == oth._get("FillChar")
            )
        if hasattr(oth, "typeName") and getattr(oth, "typeName") == "com.sun.star.style.TabStop":
            ts = cast(TabStop, oth)
            return (
                self._get("Position") == ts.Position
                and self._get("Alignment") == ts.Alignment
                and self._get("DecimalChar") == ts.DecimalChar
                and self._get("FillChar") == ts.FillChar
            )
        return NotImplemented

    # endregion dunder methods

    # region format methods
    def fmt_position(self: _TTabStopStruct, value: float | UnitT) -> _TTabStopStruct:
        """
        Gets a copy of instance with position set.

        Args:
            value (float, UnitT): Position value (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Tab: Tab instance
        """
        cp = self.copy()
        cp.prop_position = value
        return cp

    def fmt_align(self: _TTabStopStruct, value: TabAlign) -> _TTabStopStruct:
        """
        Gets a copy of instance with align set.

        Args:
            value (float): Align value.

        Returns:
            Tab: Tab instance
        """
        cp = self.copy()
        cp.prop_align = value
        return cp

    def fmt_decimal_char(self: _TTabStopStruct, value: str) -> _TTabStopStruct:
        """
        Gets a copy of instance with decimal char set.

        Args:
            value (float): Decimal char value.

        Returns:
            Tab: Tab instance
        """
        cp = self.copy()
        cp.prop_decimal_char = value
        return cp

    def fmt_fill_char(self: _TTabStopStruct, value: FillCharKind | str) -> _TTabStopStruct:
        """
        Gets a copy of instance with fill char set.

        Args:
            value (float): Fill char value.

        Returns:
            Tab: Tab instance
        """
        cp = self.copy()
        cp.prop_fill_char = value
        return cp

    # endregion format methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.STATIC
        return self._format_kind_prop

    @property
    def prop_position(self) -> UnitMM:
        """Gets/Sets the position of the tabulator in relation to the left border (in ``mm`` units)."""
        return UnitMM.from_mm100(self._get("Position"))

    @prop_position.setter
    def prop_position(self, value: float | UnitT) -> None:
        try:
            self._set("Position", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("Position", UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_align(self) -> TabAlign:
        """Gets/Sets the alignment of the text range before the tabulator"""
        return self._get("Alignment")

    @prop_align.setter
    def prop_align(self, value: TabAlign) -> None:
        self._set("Alignment", value)

    @property
    def prop_decimal_char(self) -> str:
        """Gets/Sets which delimiter is used for the decimal."""
        return self._get("DecimalChar")

    @prop_decimal_char.setter
    def prop_decimal_char(self, value: str) -> None:
        # tab align is not decimal then get_tab_stop() will replace
        # DecimalChar with " "
        # so there is no concern about letting any value being set here.
        self._set("DecimalChar", value[:1])

    @property
    def prop_fill_char(self) -> str:
        """Gets/Sets which delimiter is used for the decimal."""
        return self._get("FillChar")

    @prop_fill_char.setter
    def prop_fill_char(self, value: FillCharKind | str) -> None:
        self._set("FillChar", str(value)[:1])

    # endregion properties
