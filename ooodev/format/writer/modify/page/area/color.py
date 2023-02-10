"""
Module for Page Style Fill Color Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations

from ......exceptions import ex as mEx
from ......utils import color as mColor
from ....style.page.kind.style_page_kind import StylePageKind
from ..page_style_base import PageStyleBase


class Color(PageStyleBase):
    """
    Page Fill Coloring

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: mColor.Color = -1, style_name: StylePageKind | str = StylePageKind.STANDARD) -> None:
        """
        Constructor

        Args:
            color (Color, optional): FillColor Color.
            style_name (str, optional): Style to apply formating to. Default to the ``Default Page Style``.

        Returns:
            None:
        """
        self._style_name = str(style_name)
        init_vals = {}
        if color >= 0:
            init_vals[self._get_property_name()] = color
        else:
            init_vals[self._get_property_name()] = -1

        super().__init__(**init_vals)

    def _get_property_name(self) -> str:
        return "BackColor"

    @staticmethod
    def from_obj(obj: object, style_name: StylePageKind | str = StylePageKind.STANDARD) -> Color:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO Writer Document
            style_name (str, optional): Style to apply formating to. Default to the ``Default Page Style``.

        Raises:
            NotSupportedError: If ``obj`` is not a Writer Document.

        Returns:
            Color: Instance that represents Style Color.
        """
        bc = Color(style_name=style_name)
        if not bc._is_valid_obj(obj):
            raise mEx.NotSupportedError("obj is not a Writer Document")

        p = bc._get_style_props(obj)
        bc.prop_color = p.getPropertyValue(bc._get_property_name())
        return bc

    @property
    def prop_color(self) -> mColor.Color:
        """Gets/Sets color"""
        return self._get(self._get_property_name())

    @prop_color.setter
    def prop_color(self, value: mColor.Color):
        if value >= 0:
            self._set(self._get_property_name(), value)
        else:
            self._set(self._get_property_name(), -1)

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)
