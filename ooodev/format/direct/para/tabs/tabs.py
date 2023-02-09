"""
Modele for managing paragraph Tabs.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast

import uno

from .....events.args.cancel_event_args import CancelEventArgs
from .....meta.static_prop import static_prop
from .....utils import props as mProps
from .....utils import info as mInfo
from .....utils import lo as mLo
from ...structs.tab_stop_struct import TabStopStruct, FillCharKind as FillCharKind

from com.sun.star.beans import XPropertySet

from ooo.dyn.style.tab_align import TabAlign as TabAlign
from ooo.dyn.style.tab_stop import TabStop


class Tabs(TabStopStruct):
    """
    Paragraph Tabs

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.ParagraphProperties", "com.sun.star.style.ParagraphStyle")

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _get_property_name(self) -> str:
        return "ParaTabStops"

    @staticmethod
    def find(obj: object, position: float) -> TabStopStruct | None:
        """
        Gets a Tab that matches position from obj such as a cursor.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (float): position of tab stop (in mm units).

        Returns:
            Tab | None: ``Tab`` instance if found; Otherwise, ``None``
        """
        if not Tabs.default._is_valid_obj(obj):
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
        ts = tss[match]
        return TabStopStruct.from_tab_stop(ts)

    @classmethod
    def remove_by_pos(cls, obj: object, position: float) -> bool:
        """
        Removes a Tab Stop from ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (float): position of tab stop (in mm units).

        Returns:
            bool: ``True`` if a Tab Stop has been removed; Oherwise, ``False``
        """
        tb = cls.find(obj, position)
        if tb is None:
            return False
        # tb will contain the exact Position number so no need to plus or minus
        pos = cast(int, tb._get("Position"))
        return cls._remove_by_positon(obj, pos)

    @classmethod
    def _remove_by_positon(cls, obj: object, position: int) -> bool:
        """
        Removes a Tab Stop from ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (int): position of tab stop.

        Returns:
            None:
        """
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, Tabs.default._get_property_name()))
        if tss is None:
            return False
        lst = []
        for ts in tss:
            if position != ts.Position:
                lst.append(ts)
        if len(lst) == len(tss):
            return False
        Tabs.default._set_obj_tabs(obj, lst)
        return True

    @classmethod
    def remove(cls, obj: object, tab: TabStop | TabStopStruct) -> bool:
        """
        Removes a Tab Stop from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            tab (TabStop | Tab): Tab or Tab Stop to remove.

        Returns:
            bool: ``True`` if a Tab has been removed; Oherwise, ``False``
        """
        if not Tabs.default._is_valid_obj(obj):
            return False
        if isinstance(tab, TabStopStruct):
            return Tabs.default._remove_by_positon(obj, tab._get("Position"))
        ts = cast(TabStop, tab)
        return Tabs.default._remove_by_positon(obj, ts.Position)

    @classmethod
    def remove_all(cls, obj: object) -> None:
        """
        Removes all tab from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
        """
        if not Tabs.default._is_valid_obj(obj):
            return
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, Tabs.default._get_property_name()))
        if tss is None:
            return
        try:
            Tabs.default._set_obj_tabs(obj, [Tabs.default.get_tab_stop()])
        except Exception:
            # if for any reason can't get default it is ok to just remove all.
            Tabs.default._set_obj_tabs(obj, [])

    def get_tab_stop(self) -> TabStop:
        """
        Gets tab stop for instance

        Returns:
            TabStop: Tab stop instance
        """
        ts = super().get_tab_stop()
        try:
            # not critical
            if self is Tabs.default:
                ts.DecimalChar = "."
        except Exception:
            pass
        return ts

    # endregion methods

    # region Properties
    @static_prop
    def default() -> Tabs:  # type: ignore[misc]
        """Gets ``Tabs`` default. Static Property."""
        try:
            return Tabs._DEFAULT_INST
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
            )
            ts_val = props.getPropertyValue("TabStop")
            inst._set("Position", ts_val)
            inst._set("DecimalChar", ".")
            inst._is_default_inst = True
            Tabs._DEFAULT_INST = inst
        return Tabs._DEFAULT_INST

    # endregion Properties
