# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.utils.kind.line_spacing_mode_kind import ModeKind
from ooodev.format.inner.direct.write.para.indent_space.line_spacing import LineSpacing as InnerLineSpacing
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
# endregion Import


class LineSpacing(ParaStyleBaseMulti):
    """
    Paragraph Style Line Spacing

    .. seealso::

        - :ref:`help_writer_format_modify_para_indent_spacing`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        mode: ModeKind | None = None,
        value: int | float | UnitT = 0,
        active_ln_spacing: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            mode (ModeKind, optional): Determines the mode that is used to apply units.
            value (int, float, UnitT, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PROPORTIONAL``,
                ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            active_ln_spacing (bool, optional): Determines active page line-spacing.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        Note:
            If ``ModeKind`` is ``SINGLE``, ``LINE_1_15``, ``LINE_1_5``, or ``DOUBLE`` then ``value`` is ignored.

            If ``ModeKind`` is ``AT_LEAST``, ``LEADING``, or ``FIXED`` then ``value`` is a float (``in mm units``)
            or :ref:`proto_unit_obj`

            If ``ModeKind`` is ``PROPORTIONAL`` then value is an int representing percentage.
            For example ``95`` equals ``95%``, ``130`` equals ``130%``

        See Also:
            - :ref:`help_writer_format_modify_para_indent_spacing`
        """

        direct = InnerLineSpacing(mode=mode, value=value, active_ln_spacing=active_ln_spacing)
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
    ) -> LineSpacing:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            LineSpacing: ``LineSpacing`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerLineSpacing.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerLineSpacing:
        """Gets Inner Line Spacing instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerLineSpacing, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerLineSpacing) -> None:
        if not isinstance(value, InnerLineSpacing):
            raise TypeError(f'Expected type of InnerLineSpacing, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
