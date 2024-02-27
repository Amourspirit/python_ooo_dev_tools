# region Import
from __future__ import annotations
from typing import Any, cast
import uno

from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.inner.direct.write.para.text_flow.hyphenation import Hyphenation as InnerHyphenation

# endregion Import


class Hyphenation(ParaStyleBaseMulti):
    """
    Paragraph Style Hyphenation

    .. seealso::

        - :ref:`help_writer_format_modify_para_text_flow`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        auto: bool | None = None,
        no_caps: bool | None = None,
        start_chars: int | None = None,
        end_chars: int | None = None,
        max: int | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            auto (bool, optional): Hyphenate automatically.
            no_caps (bool, optional): Don't hyphenate word in caps.
            start_chars (int, optional): Characters at line begin.
            end_chars (int, optional): characters at line end.
            max (int, optional): Maximum consecutive hyphenated lines.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_text_flow`
        """

        direct = InnerHyphenation(auto=auto, no_caps=no_caps, start_chars=start_chars, end_chars=end_chars, max=max)
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
    ) -> Hyphenation:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Breaks: ``Breaks`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHyphenation.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerHyphenation:
        """Gets/Sets Inner Hyphenation instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerHyphenation, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerHyphenation) -> None:
        if not isinstance(value, InnerHyphenation):
            raise TypeError(f'Expected type of InnerHyphenation, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
