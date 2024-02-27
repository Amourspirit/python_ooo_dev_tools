# region Imports
from __future__ import annotations
from typing import Tuple, cast
import uno

from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind
from ooodev.format.inner.direct.structs.cell_protection_struct import CellProtectionStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti

# endregion Imports


class InnerCellProtection(CellProtectionStruct):
    """
    Style Cell Protection.

    Note:
        Cell protection is only effective after the sheet is has been applied to is also been protected.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_protection`

    .. versionadded:: 0.9.0
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


class CellProtection(CellStyleBaseMulti):
    """
    Cell Style Background Color.

    Note:
        Cell protection is only effective after the sheet is has been applied to is also been protected.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_protection`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        hide_all: bool = False,
        protected: bool = False,
        hide_formula: bool = False,
        hide_print: bool = False,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> None:
        """
        Constructor

        Args:
            hide_all (bool, optional): Specifies if all is hidden. Defaults to ``False``.
            protected (bool, optional): Specifies protected value. Defaults to ``False``.
            hide_formula (bool, optional): Specifies if the formula is hidden. Defaults to ``False``.
            hide_print (bool, optional): Specifies if the cell are to be omitted during print. Defaults to ``False``.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_protection`
        """

        direct = InnerCellProtection(
            hide_all=hide_all, protected=protected, hide_formula=hide_formula, hide_print=hide_print
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    # endregion Init

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> CellProtection:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            CellProtection: ``CellProtection`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerCellProtection.from_obj(obj=inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleCellKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerCellProtection:
        """Gets/Sets Inner Cell Protection instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerCellProtection, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerCellProtection) -> None:
        if not isinstance(value, InnerCellProtection):
            raise TypeError(f'Expected type of InnerColor, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value)

    # endregion Properties
