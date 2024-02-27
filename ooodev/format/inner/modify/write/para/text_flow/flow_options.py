# region Import
from __future__ import annotations
from typing import Any, cast
import uno

from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.format.inner.direct.write.para.text_flow.flow_options import FlowOptions as InnerFlowOptions
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti

# endregion Import


class FlowOptions(ParaStyleBaseMulti):
    """
    Paragraph Style Flow Options

    .. seealso::

        - :ref:`help_writer_format_modify_para_text_flow`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        orphans: int | None = None,
        widows: int | None = None,
        keep: bool | None = None,
        no_split: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            orphans (int, optional): Number of Orphan Control Lines.
            widows (int, optional): Number Widow Control Lines.
            keep (bool, optional): Keep with next paragraph.
            no_split (bool, optional): Do not split paragraph.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_text_flow`
        """

        direct = InnerFlowOptions(orphans=orphans, widows=widows, keep=keep, no_split=no_split)
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
    ) -> FlowOptions:
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
        direct = InnerFlowOptions.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerFlowOptions:
        """Gets/Sets Inner Flow Options instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerFlowOptions, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerFlowOptions) -> None:
        if not isinstance(value, InnerFlowOptions):
            raise TypeError(f'Expected type of InnerFlowOptions, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
