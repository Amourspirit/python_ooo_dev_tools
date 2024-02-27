# region Import
from __future__ import annotations
from typing import Tuple, cast, Type, TypeVar
import uno
from ooo.dyn.table.shadow_location import ShadowLocation

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.direct.structs.shadow_struct import ShadowStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.units.unit_obj import UnitT
from ooodev.utils.color import StandardColor, Color

# endregion Import


class InnerShadow(ShadowStruct):
    """
    Page Style Header/Footer Border Shadow

    .. seealso::

        - :ref:`help_calc_format_modify_page_header_borders`

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CellStyle",
                "com.sun.star.style.PageStyle",
            )
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "HeaderShadowFormat"
        return self._property_name

    # endregion Internal Methods

    # region properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.HEADER
        return self._format_kind_prop

    # endregion properties


_TShadow = TypeVar(name="_TShadow", bound="Shadow")


class Shadow(CellStyleBaseMulti):
    """
    Page Style Header Border Shadow

    .. seealso::

        - :ref:`help_calc_format_modify_page_header_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float | UnitT = 1.76,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow.
                Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (:py:data:`~.utils.color.Color`, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, UnitT, optional): contains the size of the shadow (in ``mm`` units)
                or :ref:`proto_unit_obj`. Defaults to ``1.76``.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_header_borders`
        """

        direct = InnerShadow(
            location=location, color=color, transparent=transparent, width=width, _cattribs=self._get_inner_cattribs()  # type: ignore
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    # region Internal Methods
    def _get_inner_prop_name(self) -> str:
        return "HeaderShadowFormat"

    def _get_inner_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_property_name": self._get_inner_prop_name(),
        }

    # endregion Internal Methods

    # region Static Methods
    @classmethod
    def from_style(
        cls: Type[_TShadow],
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> _TShadow:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Shadow: ``Shadow`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerShadow.from_obj(inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
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
    def prop_inner(self) -> InnerShadow:
        """Gets/Sets Inner Shadow instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerShadow, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerShadow) -> None:
        if not isinstance(value, InnerShadow):
            raise TypeError(f'Expected type of InnerShadow, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        cp = value.copy(_cattribs=self._get_inner_cattribs())
        self._set_style("direct", cp)

    # endregion Properties
