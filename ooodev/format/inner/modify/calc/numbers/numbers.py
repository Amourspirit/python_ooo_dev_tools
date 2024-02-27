# region Import
from __future__ import annotations
from typing import cast, Tuple, Type
import uno
from com.sun.star.lang import XComponent
from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
from ooo.dyn.lang.locale import Locale
from ooo.dyn.util.number_format import NumberFormatEnum

from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind
from ooodev.format.inner.direct.calc.numbers.numbers import Numbers as DirectNumbers
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti

# endregion Import


class InnerNumbers(DirectNumbers):
    """
    Inner Number Format

    .. seealso::

        - :ref:`help_calc_format_modify_cell_numbers`

    .. versionadded:: 0.9.4
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle",)
        return self._supported_services_values

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE
        return self._format_kind_prop


class Numbers(CellStyleBaseMulti):
    """
    Style Numbers Format.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_numbers`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        *,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
        component: XComponent | None = None,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): This property specifies the name of the font style. It may contain more than one name separated by comma.
            size (float, optional): This value contains the size of the characters in ``pt`` (point) units or :ref:`proto_unit_obj`.
            font_style (str, optional): Font style name such as ``Bold``.
            lang (Lang, optional): Font Language
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_numbers`
        """

        direct = self._get_inner_class_type()(
            num_format=num_format, num_format_index=num_format_index, lang_locale=lang_locale, component=component
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    def _get_inner_class_type(self) -> Type[DirectNumbers]:
        return InnerNumbers

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> Numbers:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            Numbers: ``Numbers`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = inst._get_inner_class_type().from_obj(inst.get_style_props(doc), component=cast(XComponent, doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleCellKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerNumbers:
        """Gets/Sets Inner Numbers instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerNumbers, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerNumbers) -> None:
        if not isinstance(value, DirectNumbers):
            raise TypeError(f'Expected type of DirectNumbers or child class, got "{type(value).__name__}"')
        inner = self.prop_inner.copy()
        inner._format_key_prop = value._format_key_prop
        self._del_attribs("_direct_inner")
        self._set_style("direct", inner)
