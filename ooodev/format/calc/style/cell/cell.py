# region Import
from __future__ import annotations
from typing import Any, Tuple

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.meta.static_prop import static_prop
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleName
from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind

# endregion Import


class Cell(StyleName):
    """
    Style Cell/Cell Range.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. seealso::

        - :ref:`help_calc_format_style_cell`

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: StyleCellKind | str = "") -> None:
        """
        Constructor

        Args:
            name (StyleCellKind, str, optional): Cell Style. Defaults to ``Default``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
        """
        if name == "":
            name = Cell.default.prop_name
        super().__init__(name=name)

    # region Overrides
    def _get_family_style_name(self) -> str:
        return "CellStyles"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.sheet.SheetCellRange",
                "com.sun.star.sheet.SheetCell",
            )
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "CellStyle"
        return self._property_name

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs):
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        if event_args.value == "":
            event_args.value = Cell.default.prop_name
        super().on_property_setting(source, event_args)

    # endregion Overrides

    # region Style Properties

    @property
    def accent_1(self) -> Cell:
        """Style Accent 1"""
        return Cell(StyleCellKind.ACCENT_1)

    @property
    def accent_2(self) -> Cell:
        """Style Accent 2"""
        return Cell(StyleCellKind.ACCENT_2)

    @property
    def accent_3(self) -> Cell:
        """Style Accent 3"""
        return Cell(StyleCellKind.ACCENT_3)

    @property
    def heading_1(self) -> Cell:
        """Style Heading 1"""
        return Cell(StyleCellKind.HEADING_1)

    @property
    def heading_2(self) -> Cell:
        """Style Heading 2"""
        return Cell(StyleCellKind.HEADING_2)

    @property
    def good(self) -> Cell:
        """Style Good"""
        return Cell(StyleCellKind.GOOD)

    @property
    def bad(self) -> Cell:
        """Style bad"""
        return Cell(StyleCellKind.BAD)

    @property
    def neutral(self) -> Cell:
        """Style Neutral"""
        return Cell(StyleCellKind.NEUTRAL)

    @property
    def error(self) -> Cell:
        """Style error"""
        return Cell(StyleCellKind.ERROR)

    @property
    def warning(self) -> Cell:
        """Style Warning"""
        return Cell(StyleCellKind.WARNING)

    @property
    def footnote(self) -> Cell:
        """Style Footnote"""
        return Cell(StyleCellKind.FOOTNOTE)

    @property
    def note(self) -> Cell:
        """Style Note"""
        return Cell(StyleCellKind.NOTE)

    @property
    def h1(self) -> Cell:
        """Style Heading 1"""
        return Cell(StyleCellKind.HEADING_1)

    @property
    def h2(self) -> Cell:
        """Style Heading 2"""
        return Cell(StyleCellKind.HEADING_2)

    # endregion Style Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.CELL
        return self._format_kind_prop

    @static_prop
    def default() -> Cell:  # type: ignore[misc]
        """Gets ``Cell`` default. Static Property."""
        try:
            return Cell._DEFAULT_INST  # type: ignore[attr-defined]
        except AttributeError:
            Cell._DEFAULT_INST = Cell(name=StyleCellKind.DEFAULT)  # type: ignore[attr-defined]
            Cell._DEFAULT_INST._is_default_inst = True  # type: ignore[attr-defined]
        return Cell._DEFAULT_INST  # type: ignore[attr-defined]
