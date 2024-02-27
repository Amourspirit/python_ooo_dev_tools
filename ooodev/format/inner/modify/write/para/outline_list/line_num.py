# region Import
from __future__ import annotations
from typing import Any, cast
from ooodev.format.inner.common.abstract.abstract_line_number import AbstractLineNumber
from ooodev.format.inner.common.abstract.abstract_line_number import LineNumberProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind

# endregion Import


class InnerLineNum(AbstractLineNumber):
    @property
    def _props(self) -> LineNumberProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = LineNumberProps(
                value="ParaLineNumberStartValue", count="ParaLineNumberCount"
            )
        return self._props_internal_attributes

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop


class LineNum(ParaStyleBaseMulti):
    """
    Paragraph Style Line Number

    .. seealso::

        - :ref:`help_writer_format_modify_para_outline_and_list`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        num_start: int = 0,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            num_start (int, optional): Restart paragraph with number.
                If ``0`` then this paragraph is include in line numbering.
                If ``-1`` then this paragraph is excluded in line numbering.
                If greater than zero this paragraph is included in line numbering and the numbering is restarted with
                value of ``num_start``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_outline_and_list`
        """

        direct = InnerLineNum(num_start=num_start)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> LineNum:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            LineNum: ``LineNum`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerLineNum.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerLineNum:
        """Gets/Sets Inner Line Number instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerLineNum, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerLineNum) -> None:
        if not isinstance(value, InnerLineNum):
            raise TypeError(f'Expected type of InnerLineNum, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
