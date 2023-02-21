"""
Modele for managing paragraph Outline and List.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, Type, cast, TypeVar, overload

from .....events.args.cancel_event_args import CancelEventArgs
from .....meta.static_prop import static_prop
from ....style_base import StyleMulti
from .....exceptions import ex as mEx
from ....kind.format_kind import FormatKind
from .outline import Outline as Outline, LevelKind as LevelKind
from .list_style import ListStyle as ParaListStyle, StyleListKind as StyleListKind
from .line_num import LineNum as LineNum

_TOutlineList = TypeVar(name="_TOutlineList", bound="OutlineList")


class OutlineList(StyleMulti):
    """
    Paragraph Outline and List

    .. versionadded:: 0.9.0
    """

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

        ls = ParaListStyle(list_style=ls_style, num_start=ls_num)
        if ls.prop_has_attribs:
            self._set_style("list_style", ls, *ls.get_attrs())
        if not ln_num is None:
            ln = LineNum(ln_num)
            self._set_style("line_num", ln, *ln.get_attrs())

    # endregion init

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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TOutlineList], obj: object) -> _TOutlineList:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TOutlineList], obj: object, **kwargs) -> _TOutlineList:
        ...

    @classmethod
    def from_obj(cls: Type[_TOutlineList], obj: object, **kwargs) -> _TOutlineList:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            OutlineList: ``OutlineList`` instance that represents ``obj`` Outline and List.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        ol = Outline.from_obj(obj)
        if ol.prop_has_attribs:
            inst._set_style("outline", ol, *ol.get_attrs())
        ls = ParaListStyle.from_obj(obj)
        if ls.prop_has_attribs:
            inst._set_style("list_style", ls, *ls.get_attrs())
        ln = LineNum.from_obj(obj)
        if ln.prop_has_attribs:
            inst._set_style("line_num", ln, *ln.get_attrs())
        return inst

    # endregion from_obj()

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    @property
    def prop_inner_outline(self) -> Outline | None:
        """Gets Outline instance"""
        try:
            return self._direct_inner_outline
        except AttributeError:
            self._direct_inner_outline = cast(Outline, self._get_style_inst("outline"))
        return self._direct_inner_outline

    @property
    def prop_inner_list_style(self) -> ParaListStyle | None:
        """Gets List Style instance"""
        try:
            return self._direct_inner_ls
        except AttributeError:
            self._direct_inner_ls = cast(ParaListStyle, self._get_style_inst("list_style"))
        return self._direct_inner_ls

    @property
    def prop_inner_line_number(self) -> LineNum | None:
        """Gets Line Number instance"""
        try:
            return self._direct_inner_ln
        except AttributeError:
            self._direct_inner_ln = cast(LineNum, self._get_style_inst("line_num"))
        return self._direct_inner_ln

    @static_prop
    def default() -> OutlineList:  # type: ignore[misc]
        """Gets ``OutlineList`` default. Static Property."""
        try:
            return OutlineList._DEFAULT_INST
        except AttributeError:
            inst = OutlineList()
            inst._set_style("outline", Outline.default, *Outline.default.get_attrs())
            inst._set_style("list_style", ParaListStyle.default, *ParaListStyle.default.get_attrs())
            inst._set_style("line_num", LineNum.default, *LineNum.default.get_attrs())
            inst._is_default_inst = True
            OutlineList._DEFAULT_INST = inst
        return OutlineList._DEFAULT_INST

    # endregion properties
