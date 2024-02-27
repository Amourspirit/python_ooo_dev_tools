# region Import
from __future__ import annotations
from typing import Tuple, cast, Type, TypeVar

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.common.abstract.abstract_padding import AbstractPadding
from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti

# endregion Import

_TPadding = TypeVar(name="_TPadding", bound="Padding")


class InnerPadding(AbstractPadding):
    """
    Page Style Footer Border Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.PageStyle",)
        return self._supported_services_values

    # endregion methods

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
                left="HeaderLeftBorderDistance",
                top="HeaderTopBorderDistance",
                right="HeaderRightBorderDistance",
                bottom="HeaderBottomBorderDistance",
            )
        return self._props_internal_attributes

    # endregion properties


class Padding(CellStyleBaseMulti):
    """
    Page Style Header Border Padding.

    .. seealso::

        - :ref:`help_calc_format_modify_page_header_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitT, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitT, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitT, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitT,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all (float, UnitT, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
                If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_header_borders`
        """

        direct = InnerPadding(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            all=padding_all,
            _cattribs=self._get_inner_cattribs(),  # type: ignore
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    # region Internal Methods
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="HeaderLeftBorderDistance",
            top="HeaderTopBorderDistance",
            right="HeaderRightBorderDistance",
            bottom="HeaderBottomBorderDistance",
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
        cls: Type[_TPadding],
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> _TPadding:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Padding: ``Padding`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPadding.from_obj(inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct)
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
    def prop_style_name(self, value: str | CalcStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerPadding:
        """Gets/Sets Inner Padding instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPadding, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerPadding) -> None:
        if not isinstance(value, InnerPadding):
            raise TypeError(f'Expected type of InnerPadding, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        cp = value.copy(_cattribs=self._get_inner_cattribs())
        self._set_style("direct", cp)

    # endregion Properties
