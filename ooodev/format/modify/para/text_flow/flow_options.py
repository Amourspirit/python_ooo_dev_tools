from __future__ import annotations
from typing import cast
import uno
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.para.text_flow.flow_options import FlowOptions as DirectFlowOptions


class FlowOptions(ParaStyleBaseMulti):
    """
    Paragraph Style Flow Options

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
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectFlowOptions(orphans=orphans, widows=widows, keep=keep, no_split=no_split)
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
    ) -> FlowOptions:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Breaks: ``Breaks`` instance from document properties.
        """
        inst = super(FlowOptions, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectFlowOptions.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectFlowOptions:
        """Gets Inner Flow Options instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectFlowOptions, self._get_style_inst("direct"))
        return self._direct_inner
