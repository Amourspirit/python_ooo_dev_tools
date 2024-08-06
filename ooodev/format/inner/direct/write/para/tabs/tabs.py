"""
Module for managing paragraph Tabs.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
import contextlib
from typing import Any, Tuple, cast, Type, TypeVar, overload

import uno
from com.sun.star.beans import XPropertySet
from ooo.dyn.style.tab_align import TabAlign
from ooo.dyn.style.tab_stop import TabStop

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.format.inner.direct.structs.tab_stop_struct import FillCharKind
from ooodev.format.inner.direct.structs.tab_stop_struct import TabStopStruct
from ooodev.loader import lo as mLo
from ooodev.meta.static_prop import static_prop
from ooodev.units.unit_obj import UnitT
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps

# endregion Import

_TTabs = TypeVar("_TTabs", bound="Tabs")


class Tabs(TabStopStruct):
    """
    Paragraph Tabs

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_tabs`

    .. versionadded:: 0.9.0
    """

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

        See Also:

            - :ref:`help_writer_format_direct_para_tabs`
        """
        super().__init__(position=position, align=align, decimal_char=decimal_char, fill_char=fill_char)

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "ParaTabStops"
        return self._property_name

    # region find()
    @overload
    @classmethod
    def find(cls: Type[_TTabs], obj: Any, position: float) -> _TTabs | None: ...

    @overload
    @classmethod
    def find(cls: Type[_TTabs], obj: Any, position: float, **kwargs) -> _TTabs | None: ...

    @classmethod
    def find(cls: Type[_TTabs], obj: Any, position: float, **kwargs) -> _TTabs | None:
        """
        Gets a Tab that matches position from obj such as a cursor.

        Args:
            obj (object): UNO Object.
            position (float): position of tab stop (in mm units).

        Returns:
            Tab | None: ``Tab`` instance if found; Otherwise, ``None``
        """
        # pylint: disable=protected-access
        # pylint: disable=no-member
        nu = cls(**kwargs)
        if not nu._is_valid_obj(obj):
            return None
        key = Tabs.default._get_property_name()
        pos = round(position * 100)
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, key))
        if tss is None:
            return None
        match = -1
        # often Writer will change Position values, 800 to 801 etc.
        # for this reason using a range to check plus or minus 2
        pos_rng = range(pos - 2, pos + 3)  # plus or minus 2
        for i, ts in enumerate(tss):
            if pos == ts.Position:
                # also covers 0
                match = i
                break
            if ts.Position in pos_rng:
                match = i
                break

        if match == -1:
            return None
        ts = tss[match]  # type: ignore
        return cls.from_uno_struct(ts, **kwargs)

    # endregion find()

    # region remove_by_pos()

    @overload
    @classmethod
    def remove_by_pos(cls, obj: Any, position: float) -> bool: ...

    @overload
    @classmethod
    def remove_by_pos(cls, obj: Any, position: float, **kwargs) -> bool: ...

    @classmethod
    def remove_by_pos(cls, obj: Any, position: float, **kwargs) -> bool:
        """
        Removes a Tab Stop from ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (float): position of tab stop (in mm units).

        Returns:
            bool: ``True`` if a Tab Stop has been removed; Otherwise, ``False``
        """
        # pylint: disable=protected-access
        tb = cls.find(obj, position, **kwargs)
        if tb is None:
            return False
        # tb will contain the exact Position number so no need to plus or minus
        pos = cast(int, tb._get("Position"))
        return cls._remove_by_position(obj, pos, **kwargs)

    @classmethod
    def _remove_by_position(cls, obj: Any, position: int, **kwargs) -> bool:
        """
        Removes a Tab Stop from ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (int): position of tab stop.

        Returns:
            None:
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, inst._get_property_name()))
        if tss is None:
            return False
        lst = [ts for ts in tss if position != ts.Position]  # type: ignore
        if len(lst) == len(tss):
            return False
        inst._set_obj_tabs(obj, lst)
        return True

    # endregion remove_by_pos()

    # region remove()
    @overload
    @classmethod
    def remove(cls, obj: Any, tab: TabStop | TabStopStruct) -> bool: ...

    @overload
    @classmethod
    def remove(cls, obj: Any, tab: TabStop | TabStopStruct, **kwargs) -> bool: ...

    @classmethod
    def remove(cls, obj: Any, tab: TabStop | TabStopStruct, **kwargs) -> bool:
        """
        Removes a Tab Stop from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            tab (TabStop | Tab): Tab or Tab Stop to remove.

        Returns:
            bool: ``True`` if a Tab has been removed; Otherwise, ``False``
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            return False
        if isinstance(tab, TabStopStruct):
            return inst._remove_by_position(obj, tab._get("Position"), **kwargs)
        ts = cast(TabStop, tab)
        return inst._remove_by_position(obj, ts.Position, **kwargs)

    # endregion remove()

    # region remove_all()
    @overload
    @classmethod
    def remove_all(cls, obj: Any) -> None: ...

    @overload
    @classmethod
    def remove_all(cls, obj: Any, **kwargs) -> None: ...

    @classmethod
    def remove_all(cls, obj: Any, **kwargs) -> None:
        """
        Removes all tab from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): UNO Object.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            return
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, inst._get_property_name()))
        if tss is None:
            return
        try:
            inst._set_obj_tabs(obj, [inst.get_uno_struct()])
        except Exception:
            # if for any reason can't get default it is ok to just remove all.
            inst._set_obj_tabs(obj, [])

    # endregion remove_all()

    def get_uno_struct(self) -> TabStop:
        """
        Gets tab stop for instance

        Returns:
            TabStop: Tab stop instance
        """
        ts = super().get_uno_struct()
        with contextlib.suppress(Exception):
            # not critical
            if self is Tabs.default:
                ts.DecimalChar = "."
        return ts

    # endregion methods

    # region Properties
    @static_prop
    def default() -> Tabs:  # type: ignore[misc]
        """Gets ``Tabs`` default. Static Property."""
        # pylint: disable=protected-access
        try:
            return Tabs._DEFAULT_INST  # type: ignore
        except AttributeError:
            inst = Tabs(align=TabAlign.DEFAULT)
            # this commented section works if not in macro mode
            # mInfo.Info.get_reg_item_prop() imports lxml
            # ts_val: str = mInfo.Info.get_reg_item_prop(
            #     "Writer/Layout/Other/TabStop", kind=mInfo.Info.RegPropKind.VALUE, idx=1
            # )
            # inst._set("Position", int(ts_val))
            props = mLo.Lo.qi(
                XPropertySet,
                mInfo.Info.get_config(node_str="Other", node_path="/org.openoffice.Office.Writer/Layout/"),
                True,
            )
            ts_val = props.getPropertyValue("TabStop")
            inst._set("Position", ts_val)
            inst._set("DecimalChar", ".")
            inst._is_default_inst = True
            Tabs._DEFAULT_INST = inst  # type: ignore
        return Tabs._DEFAULT_INST  # type: ignore

    # endregion Properties
