from __future__ import annotations
from typing import Tuple, cast
import uno
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum

from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ...page_style_base_multi import PageStyleBaseMulti
from .....direct.structs.side import Side as Side, LineSize as LineSize, SideFlags as SideFlags
from .....direct.common.abstract.abstract_sides import AbstractSides
from .....direct.common.props.border_props import BorderProps
from .....kind.format_kind import FormatKind


class FooterSides(AbstractSides):
    """
    Page Footer Style Border Sides.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Sides properties.

    .. versionadded:: 0.9.0
    """

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.style.ParagraphStyle",
            "com.sun.star.style.PageStyle",
        )

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.DOC | FormatKind.STYLE

    @property
    def _props(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="FooterLeftBorder", top="FooterTopBorder", right="FooterRightBorder", bottom="FooterBottomBorder"
            )
        return self.__border_properties

    # endregion Properties


class Sides(PageStyleBaseMulti):
    """
    Page Footer Style Border Sides.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = FooterSides(left=left, right=right, top=top, bottom=bottom, border_side=border_side)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Sides:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Sides: ``Sides`` instance from document properties.
        """
        inst = super(Sides, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = FooterSides.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> FooterSides:
        """Gets Inner Sides instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(FooterSides, self._get_style_inst("direct"))
        return self._direct_inner
