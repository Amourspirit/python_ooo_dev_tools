# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind
from ooodev.format.inner.direct.write.para.drop_cap.drop_caps import DropCaps as InnerDropCaps
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
# endregion Import


class DropCaps(ParaStyleBaseMulti):
    """
    Paragraph Style Drop Caps

    .. seealso::

        - :ref:`help_writer_format_modify_para_drop_caps`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        count: int = 0,
        spaces: float | UnitT = 0.0,
        lines: int = 3,
        style: StyleCharKind | str | None = None,
        whole_word: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            count (int): Specifies the number of characters in the drop cap. Must be from ``0`` to ``255``.
            spaces (float, UnitT): Specifies the distance between the drop cap in the following text
                (in ``mm`` units) or :ref:`proto_unit_obj`.
            lines (int): Specifies the number of lines used for a drop cap. Must be from ``0`` to ``255``.
            style (StyleCharKind, str, optional): Specifies the character style name for drop caps.
            whole_word (bool, optional): specifies if Drop Cap is applied to the whole first word.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_drop_caps`
        """

        direct = InnerDropCaps(count=count, spaces=spaces, lines=lines, style=style, whole_word=whole_word)
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
    ) -> DropCaps:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            DropCaps: ``DropCaps`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerDropCaps.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerDropCaps:
        """Gets/Sets Inner Drop Caps instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerDropCaps, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerDropCaps) -> None:
        if not isinstance(value, InnerDropCaps):
            raise TypeError(f'Expected type of InnerDropCaps, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
