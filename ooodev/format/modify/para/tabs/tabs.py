from __future__ import annotations
from typing import cast
import uno
from ooo.dyn.style.tab_align import TabAlign as TabAlign
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.para.tabs.tabs import Tabs as DirectTabs
from ....direct.structs.tab_stop_struct import FillCharKind as FillCharKind


class Tabs(ParaStyleBaseMulti):
    """
    Paragraph Style Breaks

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        position: float = 0.0,
        align: TabAlign = TabAlign.LEFT,
        decimal_char: str = ".",
        fill_char: FillCharKind | str = FillCharKind.NONE,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            position (float): Specifies the position of the tabulator in relation to the left border (in mm units).
                Defaults to ``0.0``
            align (TabAlign): Specifies the alignment of the text range before the tabulator. Defaults to ``TabAlign.LEFT``
            decimal_char (str): Specifies which delimiter is used for the decimal.
                Argument is expected to be a single character string.
                This argument is only used when ``align`` is set to ``TabAlign.DECIMAL``.
            fill_char (FillCharKind, str): specifies the character that is used to fill up the space between the text in the text range and the tabulators.
                If string value then argument is expected to be a single character string.
                Defaults to ``FillCharKind.NONE``
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectTabs(position=position, align=align, decimal_char=decimal_char, fill_char=fill_char)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        index: int = 0,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Tabs:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            index (int, optional): Index of tab stop. Defatult ``0``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Tabs: ``Tabs`` instance from document properties.
        """
        inst = super(Tabs, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectTabs.from_obj(obj=inst.get_style_props(doc), index=index)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def find(
        cls,
        doc: object,
        position: float,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Tabs | None:
        """
        _summary_

        Args:
            doc (object): UNO Documnet Object.
            position (float): position of tab stop (in mm units).
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Tabs | None: ``Tab`` instance if found; Otherwise, ``None``
        """
        inst = super(Tabs, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)

        direct = DirectTabs.find(obj=inst.get_style_props(doc), position=position)
        if direct is None:
            return None
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
    def prop_inner(self) -> DirectTabs:
        """Gets Inner Tabs instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectTabs, self._get_style_inst("direct"))
        return self._direct_inner