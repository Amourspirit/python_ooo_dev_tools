# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ooodev.format.inner.direct.write.para.indent_space.spacing import Spacing as InnerSpacing
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
# endregion Import


class Spacing(ParaStyleBaseMulti):
    """
    Paragraph Style Spacing

    .. seealso::

        - :ref:`help_writer_format_modify_para_indent_spacing`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        above: float | UnitT | None = None,
        below: float | UnitT | None = None,
        style_no_space: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            above (float, UnitT, optional): Determines the top margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            below (float, UnitT, optional): Determines the bottom margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            style_no_space (bool, optional): Do not add space between paragraphs of the same style.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_indent_spacing`
        """

        direct = InnerSpacing(above=above, below=below, style_no_space=style_no_space)
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
    ) -> Spacing:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Spacing: ``Spacing`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerSpacing.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerSpacing:
        """Gets Inner Spacing instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerSpacing, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerSpacing) -> None:
        if not isinstance(value, InnerSpacing):
            raise TypeError(f'Expected type of InnerSpacing, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
