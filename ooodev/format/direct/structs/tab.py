"""
Modele for addin paragraph tab.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Dict, Iterable, Tuple, cast, overload
from enum import Enum

import uno
from ....events.event_singleton import _Events
from ....exceptions import ex as mEx
from ....utils import info as mInfo
from ....utils import lo as mLo
from ....utils import props as mProps
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent

from com.sun.star.beans import XPropertySet

from ooo.dyn.style.tab_align import TabAlign as TabAlign
from ooo.dyn.style.tab_stop import TabStop


class FillCharKind(Enum):
    """Tab Fill Char"""

    NONE = " "
    DECIMAL = "."
    DASH = "-"
    UNDER_SCORE = "_"

    def __str__(self) -> str:
        return self.value


class Tab(StyleBase):
    """
    Paragraph Tab

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        position: float = 0.0,
        align: TabAlign = TabAlign.LEFT,
        decimal_char: str = ".",
        fill_char: FillCharKind | str = FillCharKind.NONE,
    ) -> None:
        """
        Constructor

        Args:
            position (float): Specifies the position of the tabulator in relation to the left border (in mm units).
                Defaults to ``0.0``
            align (TabAlign): Specifies the alignment of the text range before the tabulator. Defaults to ``TabAlign.LEFT``
            decimal_char (str): Specifies which delimiter is used for the decimal.
                Argument is expected to be a single character string.
                This argument is only used when ``align`` is set to ``TabAlign.DECIMAL``.
            fill_char (FillCharKind, str): specifies the character that is used to fill up the space between the text in the text range and the tabulators.
                If string value then argument is expected to be a single character string.
                Defaults to ``FillCharKind.NONE``

        Returns:
            None:

        Note:
            If argument ``type`` is ``None`` then all other argument are ignored
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        # in Writer when a tab is added it is static. That is it remains when subquent paragraphs are added.
        # existing paragraph can can add or remove Tabs. This will not affect the tabs of other existing paragraphs.
        # most likely it would be best practice not to backup and restore this property (ParaLineSpacing) when appending paragraphs in Writer.

        init_vals = {"FillChar": str(fill_char)[:1], "Alignment": align, "DecimalChar": decimal_char[:1]}
        init_vals["Position"] = round(position * 100)
        if align != TabAlign.DECIMAL:
            init_vals["DecimalChar"] = " "

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

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

        If a Tab instance with the same position is existing it is updated;
        Otherwise, a new Tab is added.

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
            keys (Dict[str, str], optional): Property key, value items that map properties.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(self._supported_services()[0])

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        keys = {"prop": "ParaTabStops"}
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

        ts = self.get_tab_stop()
        tss_lst = list(tss)
        if match >= 0:
            tss_lst[match] = ts
        else:
            tss_lst.append(ts)

        Tab._set_obj_tabs(obj, tss_lst, key)
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

        # mProps.Props.set(obj, **{key: tuple(tss_lst)})

    # endregion apply()

    def get_attrs(self) -> Tuple[str, ...]:
        return ("ParaTabStops",)

    def get_tab_stop(self) -> TabStop:
        """
        Gets tab stop for instance

        Returns:
            TabStop: Tab stop instance
        """
        if self.prop_align == TabAlign.DECIMAL:
            dec = self._get("DecimalChar")
        else:
            dec = " "
        return TabStop(
            Position=self._get("Position"),
            Alignment=self._get("Alignment"),
            DecimalChar=dec,
            FillChar=self._get("FillChar"),
        )

    @classmethod
    def from_obj(cls, obj: object, index: int = 0) -> Tab:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            Tab: ``Tab`` instance that represents ``obj`` Tab properties.
        """
        if not mInfo.Info.support_service(obj, "com.sun.star.style.ParagraphProperties"):
            raise mEx.NotSupportedServiceError("com.sun.star.style.ParagraphProperties")

        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, "ParaTabStops"))
        # can expcet for ts to contain at least one TabAlign
        ts = tss[index]

        return Tab.from_tab_stop(ts)

    @classmethod
    def from_tab_stop(cls, ts: TabStop) -> Tab:
        """
        Converts a Tab Stop instance to a Tab

        Args:
            ts (TabStop): Tab stop

        Returns:
            Tab: Tab set with Tab Stop properties
        """
        inst = Tab()
        inst._set("FillChar", ts.FillChar)
        inst._set("Alignment", ts.Alignment)
        inst._set("DecimalChar", ts.DecimalChar)
        inst._set("Position", ts.Position)
        return inst

    @staticmethod
    def _set_obj_tabs(obj: object, tabs: Iterable[TabStop], prop: str = "ParaTabStops") -> None:
        prop_set = mLo.Lo.qi(XPropertySet, obj, raise_err=True)
        seq = uno.Any("[]com.sun.star.style.TabStop", tabs)
        uno.invoke(prop_set, "setPropertyValue", (prop, seq))

    # endregion methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Tab):
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
    def fmt_position(self, value: float) -> Tab:
        """
        Gets a copy of instance with position set.

        Args:
            value (float): Position value.

        Returns:
            Tab: Tab instance
        """
        cp = self.copy()
        cp.prop_position = value
        return cp

    def fmt_align(self, value: TabAlign) -> Tab:
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

    def fmt_decimal_char(self, value: str) -> Tab:
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

    def fmt_fill_char(self, value: FillCharKind | str) -> Tab:
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
        return FormatKind.PARA | FormatKind.STATIC

    @property
    def prop_position(self) -> float:
        """Gets/Sets the position of the tabulator in relation to the left border (in mm units)."""
        pv = self._get("Position")
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_position.setter
    def prop_position(self, value: float) -> None:
        self._set("Position", round(value * 100))

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
