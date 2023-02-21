from __future__ import annotations
from typing import Type, cast
import uno
from numbers import Real
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.structs.line_spacing_struct import ModeKind as ModeKind
from ....direct.para.indent_space.line_spacing import LineSpacing as DirectLineSpacing


class LineSpacing(ParaStyleBaseMulti):
    """
    Paragraph Style Line Spacing

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        mode: ModeKind | None = None,
        value: Real = 0,
        active_ln_spacing: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            mode (ModeKind, optional): Determines the mode that is use to apply units.
            value (Real, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PROPORTIONAL``, ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            active_ln_spacing (bool, optional): Determines active page line-spacing.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectLineSpacing(mode=mode, value=value, active_ln_spacing=active_ln_spacing)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> LineSpacing:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            LineSpacing: ``LineSpacing`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = DirectLineSpacing.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectLineSpacing:
        """Gets Inner Line Spacing instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectLineSpacing, self._get_style_inst("direct"))
        return self._direct_inner
