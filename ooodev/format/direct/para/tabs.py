"""
Modele for managing paragraph Tabs.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast

import uno

from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ..structs.tab import Tab, FillChar as FillChar

from ooo.dyn.style.tab_align import TabAlign as TabAlign
from ooo.dyn.style.tab_stop import TabStop


class Tabs(Tab):
    """
    Paragraph Tabs

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region methods

    @staticmethod
    def find(obj: object, position: float) -> Tab | None:
        """
        Gets a Tab that matches position from obj such as a cursor.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (float): position of tab stop (in mm units).

        Returns:
            Tab | None: ``Tab`` instance if found; Otherwise, ``None``
        """
        if not Tabs.default._is_valid_service(obj):
            return None
        key = "ParaTabStops"
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
        return Tab.from_tab_stop(ts)

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
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, "ParaTabStops"))
        if tss is None:
            return False
        lst = []
        for ts in tss:
            if position != ts.Position:
                lst.append(ts)
        if len(lst) == len(tss):
            return False
        cls._set_obj_tabs(obj, lst)
        return True

    @classmethod
    def remove(cls, obj: object, tab: TabStop | Tab) -> bool:
        """
        Removes a Tab Stop from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            tab (TabStop | Tab): Tab or Tab Stop to remove.

        Returns:
            bool: ``True`` if a Tab has been removed; Oherwise, ``False``
        """
        if not Tabs.default._is_valid_service(obj):
            return False
        if isinstance(tab, Tab):
            return cls._remove_by_positon(obj, tab._get("Position"))
        ts = cast(TabStop, tab)
        return cls._remove_by_positon(obj, ts.Position)

    @classmethod
    def remove_all(cls, obj: object) -> None:
        """
        Removes all tab from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
        """
        if not Tabs.default._is_valid_service(obj):
            return
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, "ParaTabStops"))
        if tss is None:
            return
        cls._set_obj_tabs(obj, [Tabs.default.get_tab_stop()])

    # endregion methods

    # region Properties
    @static_prop
    def default() -> Tabs:  # type: ignore[misc]
        """Gets ``Tabs`` default. Static Property."""
        if Tabs._DEFAULT is None:
            inst = Tabs(align=TabAlign.DEFAULT)
            inst._set("Position", 1251)
            Tabs._DEFAULT = inst
        return Tabs._DEFAULT

    # endregion Properties
