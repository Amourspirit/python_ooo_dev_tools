"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import cast

import uno

from ......exceptions import ex as mEx
from ......utils import lo as mLo
from ......utils import props as mProps
from .....direct.common.abstract_sides import AbstractSides, BorderProps
from .....direct.structs.side import Side as Side, BorderLineStyleEnum as BorderLineStyleEnum
from ....style.page.kind import StylePageKind as StylePageKind
from ..page_style_base import PageStyleBase

from ooo.dyn.table.border_line2 import BorderLine2

# endregion imports


class Borders(PageStyleBase, AbstractSides):
    """
    Page Style Border.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Sides properties.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
    ) -> None:
        self._style_name = str(style_name)
        super().__init__(left=left, right=right, top=top, bottom=bottom, border_side=border_side)

    # region methods
    def _is_valid_obj(self, obj: object) -> bool:
        valid = PageStyleBase._is_valid_obj(self, obj)
        if valid is False:
            valid = mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
        return valid

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies Style to obj

        Args:
            obj (object): UNO object

        Returns:
            None:
        """
        try:
            p = self._get_style_props(obj)
            AbstractSides.apply(self, p)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @staticmethod
    def from_obj(obj: object, style_name: StylePageKind | str = StylePageKind.STANDARD) -> Borders:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO Writer Document
            style_name (str, optional): Style to apply formating to. Default to the ``Default Page Style``.

        Raises:
            NotSupportedError: If ``obj`` is not a Writer Document.

        Returns:
            Sides: Instance that represents ``BorderLine2``.
        """
        bc = Borders(style_name=style_name)
        if not bc._is_valid_obj(obj):
            raise mEx.NotSupportedError("obj is not a Writer Document")

        empty = BorderLine2()
        p = bc._get_style_props(obj)
        for attr in bc._props:
            b2 = cast(BorderLine2, mProps.Props.get(p, attr, empty))
            side = Side.from_border2(b2)
            bc._set(attr, side)
        return bc

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Borders):
            eq = AbstractSides.__eq__(self, oth)
            if eq:
                eq = eq and (self.prop_style_name == oth.prop_style_name)
            return eq
        return NotImplemented

    # endregion methods

    # region Properties

    @property
    def _props(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="LeftBorder", top="TopBorder", right="RightBorder", bottom="BottomBorder"
            )
        return self.__border_properties

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    # endregion Properties
