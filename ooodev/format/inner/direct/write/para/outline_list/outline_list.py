"""
Module for managing paragraph Outline and List.

.. seealso::

    - :ref:`help_writer_format_direct_para_outline_and_list`

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, Type, cast, TypeVar, overload

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.write.para.outline_list.line_num import LineNum
from ooodev.format.inner.direct.write.para.outline_list.list_style import ListStyle as ParaListStyle
from ooodev.format.inner.direct.write.para.outline_list.outline import LevelKind
from ooodev.format.inner.direct.write.para.outline_list.outline import Outline as InnerOutline
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.writer.style.lst.style_list_kind import StyleListKind

_TOutlineList = TypeVar(name="_TOutlineList", bound="OutlineList")


class OutlineList(StyleMulti):
    """
    Paragraph Outline and List

    .. seealso::

        - :ref:`help_writer_format_direct_para_outline_and_list`

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
                Otherwise, number starts at the value passed in.
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

        See Also:

            - :ref:`help_writer_format_direct_para_outline_and_list`
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        super().__init__()
        if ol_level is not None:
            ol = InnerOutline(level=ol_level)
            self._set_style("outline", ol, *ol.get_attrs())  # type: ignore

        ls = ParaListStyle(list_style=ls_style, num_start=ls_num)
        if ls.prop_has_attribs:
            self._set_style("list_style", ls, *ls.get_attrs())  # type: ignore
        if ln_num is not None:
            ln = LineNum(ln_num)
            self._set_style("line_num", ln, *ln.get_attrs())  # type: ignore

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

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TOutlineList], obj: Any) -> _TOutlineList: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TOutlineList], obj: Any, **kwargs) -> _TOutlineList: ...

    @classmethod
    def from_obj(cls: Type[_TOutlineList], obj: Any, **kwargs) -> _TOutlineList:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            OutlineList: ``OutlineList`` instance that represents ``obj`` Outline and List.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        ol = InnerOutline.from_obj(obj)
        if ol.prop_has_attribs:
            inst._set_style("outline", ol, *ol.get_attrs())
        ls = ParaListStyle.from_obj(obj)
        if ls.prop_has_attribs:
            inst._set_style("list_style", ls, *ls.get_attrs())
        ln = LineNum.from_obj(obj)
        if ln.prop_has_attribs:
            inst._set_style("line_num", ln, *ln.get_attrs())
        inst.set_update_obj(obj)
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
    def prop_inner_outline(self) -> InnerOutline | None:
        """Gets Outline instance"""
        try:
            return self._direct_inner_outline
        except AttributeError:
            self._direct_inner_outline = cast(InnerOutline, self._get_style_inst("outline"))
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

    @property
    def default(self: _TOutlineList) -> _TOutlineList:
        """Gets ``OutlineList`` default."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            if self.prop_inner_line_number is None:
                ln = LineNum().default
            else:
                ln = self.prop_inner_line_number.default
            if self.prop_inner_outline is None:
                ol = InnerOutline().default
            else:
                ol = self.prop_inner_outline.default
            if self.prop_inner_list_style is None:
                ls = ParaListStyle().default
            else:
                ls = self.prop_inner_list_style.default
            inst = self.__class__(_cattribs=self._get_internal_cattribs())
            inst._set_style("outline", ol, *ol.get_attrs())  # type: ignore
            inst._set_style("list_style", ls, *ls.get_attrs())  # type: ignore
            inst._set_style("line_num", ln, *ln.get_attrs())  # type: ignore
            inst._is_default_inst = True
            self._default_inst = inst
        return self._default_inst

    # endregion properties
