# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooo.dyn.style.tab_align import TabAlign
from ooodev.format.inner.direct.structs.tab_stop_struct import FillCharKind
from ooodev.format.inner.direct.write.para.tabs.tabs import Tabs as InnerTabs
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import


class Tabs(ParaStyleBaseMulti):
    """
    Paragraph Style Breaks

    .. seealso::

        - :ref:`help_writer_format_modify_para_tabs`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        position: float | UnitT = 0.0,
        align: TabAlign = TabAlign.LEFT,
        decimal_char: str = ".",
        fill_char: FillCharKind | str = FillCharKind.NONE,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            position (float, UnitT, optional): Specifies the position of the tabulator in relation to the left border
                (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0.0``
            align (TabAlign, optional): Specifies the alignment of the text range before the tabulator. Defaults to
                ``TabAlign.LEFT``
            decimal_char (str, optional): Specifies which delimiter is used for the decimal.
                Argument is expected to be a single character string.
                This argument is only used when ``align`` is set to ``TabAlign.DECIMAL``.
            fill_char (FillCharKind, str, optional): specifies the character that is used to fill up the space between
                the text in the text range and the tabulators.
                If string value then argument is expected to be a single character string.
                Defaults to ``FillCharKind.NONE``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_tabs`
        """

        direct = InnerTabs(position=position, align=align, decimal_char=decimal_char, fill_char=fill_char)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: Any,
        index: int = 0,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Tabs:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            index (int, optional): Index of tab stop. Default ``0``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Tabs: ``Tabs`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerTabs.from_obj(obj=inst.get_style_props(doc), index=index)
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
            doc (object): UNO Document Object.
            position (float): position of tab stop (in mm units).
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Tabs | None: ``Tab`` instance if found; Otherwise, ``None``
        """
        inst = cls(style_name=style_name, style_family=style_family)

        direct = InnerTabs.find(obj=inst.get_style_props(doc), position=position)
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
    def prop_inner(self) -> InnerTabs:
        """Gets/Sets Inner Tabs instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerTabs, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerTabs) -> None:
        if not isinstance(value, InnerTabs):
            raise TypeError(f'Expected type of InnerTabs, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
