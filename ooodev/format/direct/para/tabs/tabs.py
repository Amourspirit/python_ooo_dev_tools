"""
Modele for managing paragraph Tabs.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, Type, TypeVar, overload

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

_TTabs = TypeVar(name="_TTabs", bound="Tabs")


class Tabs(TabStopStruct):
    """
    Paragraph Tabs

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

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

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "ParaTabStops"
        return self._property_name

    # region find()
    @overload
    @classmethod
    def find(cls: Type[_TTabs], obj: object, position: float) -> _TTabs | None:
        ...

    @overload
    @classmethod
    def find(cls: Type[_TTabs], obj: object, position: float, **kwargs) -> _TTabs | None:
        ...

    @classmethod
    def find(cls: Type[_TTabs], obj: object, position: float, **kwargs) -> _TTabs | None:
        """
        Gets a Tab that matches position from obj such as a cursor.

        Args:
            obj (object): UNO Object.
            position (float): position of tab stop (in mm units).

        Returns:
            Tab | None: ``Tab`` instance if found; Otherwise, ``None``
        """
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
        ts = tss[match]
        return cls.from_uno_struct(ts, **kwargs)

    # endregion find()

    # region remove_by_pos()

    @overload
    @classmethod
    def remove_by_pos(cls, obj: object, position: float) -> bool:
        ...

    @overload
    @classmethod
    def remove_by_pos(cls, obj: object, position: float, **kwargs) -> bool:
        ...

    @classmethod
    def remove_by_pos(cls, obj: object, position: float, **kwargs) -> bool:
        """
        Removes a Tab Stop from ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (float): position of tab stop (in mm units).

        Returns:
            bool: ``True`` if a Tab Stop has been removed; Oherwise, ``False``
        """
        tb = cls.find(obj, position, **kwargs)
        if tb is None:
            return False
        # tb will contain the exact Position number so no need to plus or minus
        pos = cast(int, tb._get("Position"))
        return cls._remove_by_positon(obj, pos, **kwargs)

    @classmethod
    def _remove_by_positon(cls, obj: object, position: int, **kwargs) -> bool:
        """
        Removes a Tab Stop from ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            position (int): position of tab stop.

        Returns:
            None:
        """
        inst = cls(**kwargs)
        tss = cast(Tuple[TabStop, ...], mProps.Props.get(obj, inst._get_property_name()))
        if tss is None:
            return False
        lst = []
        for ts in tss:
            if position != ts.Position:
                lst.append(ts)
        if len(lst) == len(tss):
            return False
        inst._set_obj_tabs(obj, lst)
        return True

    # endregion remove_by_pos()

    # region remove()
    @overload
    @classmethod
    def remove(cls, obj: object, tab: TabStop | TabStopStruct) -> bool:
        ...

    @overload
    @classmethod
    def remove(cls, obj: object, tab: TabStop | TabStopStruct, **kwargs) -> bool:
        ...

    @classmethod
    def remove(cls, obj: object, tab: TabStop | TabStopStruct, **kwargs) -> bool:
        """
        Removes a Tab Stop from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties``
            tab (TabStop | Tab): Tab or Tab Stop to remove.

        Returns:
            bool: ``True`` if a Tab has been removed; Oherwise, ``False``
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            return False
        if isinstance(tab, TabStopStruct):
            return inst._remove_by_positon(obj, tab._get("Position"), **kwargs)
        ts = cast(TabStop, tab)
        return inst._remove_by_positon(obj, ts.Position, **kwargs)

    # endregion remove()

    # region remove_all()
    @overload
    @classmethod
    def remove_all(cls, obj: object) -> None:
        ...

    @overload
    @classmethod
    def remove_all(cls, obj: object, **kwargs) -> None:
        ...

    @classmethod
    def remove_all(cls, obj: object, **kwargs) -> None:
        """
        Removes all tab from ``obj`` ``ParaTabStops`` property.

        Args:
            obj (object): UNO Object.
        """
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
