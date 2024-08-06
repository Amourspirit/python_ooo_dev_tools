# region Import
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar
import uno

from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.direct.structs.side import Side
from ooodev.format.inner.direct.write.char.border.sides import Sides as DirectSides
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.page.page_style_base_multi import PageStyleBaseMulti
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind

# endregion Import


class InnerSides(DirectSides):
    """
    Page Header/Footer Style Border Sides.

    .. versionadded:: 0.9.0
    """

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.PageStyle",)
        return self._supported_services_values

    # endregion overrides

    # region properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.HEADER
        return self._format_kind_prop

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="HeaderLeftBorder", top="HeaderTopBorder", right="HeaderRightBorder", bottom="HeaderBottomBorder"
            )
        return self._props_internal_attributes

    # endregion properties


_TSides = TypeVar("_TSides", bound="Sides")


class Sides(PageStyleBaseMulti):
    """
    Page Header Style Border Sides.

    .. seealso::

        - :ref:`help_writer_format_modify_page_header_borders`

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
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            all (Side | None, optional): Determines the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            style_name (WriterStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_header_borders`
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
        doc: Any,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> _TSides:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (WriterStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

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
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.HEADER
        return self._format_kind_prop

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
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
        cp = value.copy(_cattribs=self._get_inner_cattribs())
        self._set_style("direct", cp, *cp.get_attrs())

    # endregion Properties
