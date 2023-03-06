from __future__ import annotations
from typing import Tuple, cast, Type, TypeVar
import uno
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum

from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ...page_style_base_multi import PageStyleBaseMulti
from .....direct.structs.side import Side as Side, LineSize as LineSize
from .....direct.char.border.sides import Sides as InnerSides
from .....direct.common.props.border_props import BorderProps

_TSides = TypeVar(name="_TSides", bound="Sides")


class Sides(PageStyleBaseMulti):
    """
    Page Header Style Border Sides.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        all: Side | None = None,
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
            all (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignoreded
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = InnerSides(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            all=all,
            _cattribs=self._get_inner_cattribs(),
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    # region Internal Methods
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="HeaderLeftBorder", top="HeaderTopBorder", right="HeaderRightBorder", bottom="HeaderBottomBorder"
        )

    def _get_inner_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_props_internal_attributes": self._get_inner_props(),
        }

    # endregion Internal Methods

    # region Static Methods
    @classmethod
    def from_style(
        cls: Type[_TSides],
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> _TSides:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Sides: ``Sides`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerSides.from_obj(inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerSides:
        """Gets/Sets Inner Sides instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerSides, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerSides) -> None:
        if not isinstance(value, InnerSides):
            raise TypeError(f'Expected type of InnerSides, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())

    # endregion Properties
