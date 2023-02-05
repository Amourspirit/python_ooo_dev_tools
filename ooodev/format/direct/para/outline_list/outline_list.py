"""
Modele for managing paragraph Outline and List.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple

from .....meta.static_prop import static_prop
from ....style_base import StyleMulti
from ....kind.format_kind import FormatKind
from .outline import Outline as Outline, LevelKind as LevelKind
from .list_style import ListStyle as ListStyle, StyleListKind as StyleListKind
from .line_num import LineNum as LineNum


class OutlineList(StyleMulti):
    """
    Paragraph Outline and List

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        *,
        ol_level: LevelKind | None = None,
        ls_style: str | StyleListKind | None = None,
        ls_num: int | None = None,
        ln_num: int | None = None,
    ) -> None:
        """
        Constructor

        Args:
            ol_level (LevelKind, optional): Outline level.
            ls_style (str, StyleListKind, optional): List Style.
            ls_num (int, optional): Starts with number.
                If ``-1`` then restart numbering at current paragraph is consider to be ``False``.
                If ``-2`` then restart numbering at current paragraph is consider to be ``True``.
                Otherewise, restart numbering is considered to be ``True``.
            ln_num (int, optional): Restart paragraph with number.
                If ``0`` then this paragraph is include in line numbering.
                If ``-1`` then this paragraph is excluded in line numbering.
                If greater then zero then this paragraph is included in line numbering and the numbering is restarted with value of ``ln_num``.
        Returns:
            None:

        Note:
            Arguments that start with ``ol_`` set Outline.

            Arguments that start with ``ls_`` set List Style.

            Arguments that start with ``ln_`` set Line Spacing.
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        super().__init__()
        if not ol_level is None:
            ol = Outline(level=ol_level)
            self._set_style("outline", ol, *ol.get_attrs())

        ls = ListStyle(list_style=ls_style, num_start=ls_num)
        if ls.prop_has_attribs:
            self._set_style("list_style", ls, *ls.get_attrs())
        if not ln_num is None:
            ln = LineNum(ln_num)
            self._set_style("line_num", ln, *ln.get_attrs())

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    @staticmethod
    def from_obj(obj: object) -> OutlineList:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            OutlineList: ``OutlineList`` instance that represents ``obj`` Outline and List.
        """
        inst = OutlineList()
        ol = Outline.from_obj(obj)
        if ol.prop_has_attribs:
            inst._set_style("outline", ol, *ol.get_attrs())
        ls = ListStyle.from_obj(obj)
        if ls.prop_has_attribs:
            inst._set_style("list_style", ls, *ls.get_attrs())
        ln = LineNum.from_obj(obj)
        if ln.prop_has_attribs:
            inst._set_style("line_num", ln, *ln.get_attrs())
        return inst

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @static_prop
    def default() -> OutlineList:  # type: ignore[misc]
        """Gets ``OutlineList`` default. Static Property."""
        if OutlineList._DEFAULT is None:
            inst = OutlineList()
            inst._set_style("outline", Outline.default.copy(), *Outline.default.get_attrs())
            inst._set_style("list_style", ListStyle.default.copy(), *ListStyle.default.get_attrs())
            inst._set_style("line_num", LineNum.default.copy(), *LineNum.default.get_attrs())
            OutlineList._DEFAULT = inst
        return OutlineList._DEFAULT

    # endregion properties
