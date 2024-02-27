# region Imports
from __future__ import annotations
from typing import NamedTuple
import uno

from ooodev.utils import props as mProps
from ooodev.exceptions import ex as mEx
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.modify.calc.cell_style_base import CellStyleBase

# endregion Imports


class PrintProps(NamedTuple):
    """Inner Properties"""

    headers: str
    grid: str
    comments: str
    obj_img: str
    charts: str
    draw_obj: str
    formula: str
    zero_val: str


class Printing(CellStyleBase):
    """
    Page Style Order.

    .. seealso::

        - :ref:`help_calc_format_modify_page_sheet`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        header: bool | None = None,
        grid: bool | None = None,
        comment: bool | None = None,
        obj_img: bool | None = None,
        chart: bool | None = None,
        drawing: bool | None = None,
        formula: bool | None = None,
        zero_value: bool | None = None,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            header (bool, optional): Specifies if headers are printed.
            grid (bool, optional): Specifies if grids are printed.
            comment (bool, optional): Specifies if comments are printed.
            obj_img (bool, optional): Specifies if objects and images are printed.
            chart (bool, optional): Specifies if charts are printed.
            drawing (bool, optional): Specifies if drawings are printed.
            formula (bool, optional): Specifies if formulas are printed.
            zero_value (bool, optional): Specifies if zero values are printed.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_sheet`
        """

        super().__init__(style_name=style_name, style_family=style_family)
        if header is not None:
            self.prop_header = header
        if grid is not None:
            self.prop_grid = grid
        if comment is not None:
            self.prop_comment = comment
        if obj_img is not None:
            self.prop_obj_img = obj_img
        if chart is not None:
            self.prop_chart = chart
        if drawing is not None:
            self.prop_drawing = drawing
        if formula is not None:
            self.prop_formula = formula
        if zero_value is not None:
            self.prop_zero_value = zero_value

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> Printing:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Printing: ``Printing`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        style_props = inst.get_style_props(doc)
        if not inst._is_valid_obj(style_props):
            raise mEx.NotSupportedError(f"Object is not support to convert to {cls.__name__}")

        for prop in inst._props:
            if prop:
                inst._set(prop, mProps.Props.get(style_props, prop))

        return inst

    # region Properties
    @property
    def prop_header(self) -> bool | None:
        """
        Gets/Sets Print Headers
        """
        return self._get(self._props.headers)

    @prop_header.setter
    def prop_header(self, value: bool | None):
        if value is None:
            self._remove(self._props.headers)
            return
        self._set(self._props.headers, value)

    @property
    def prop_grid(self) -> bool | None:
        """
        Gets/Sets Print Grid.
        """
        return self._get(self._props.grid)

    @prop_grid.setter
    def prop_grid(self, value: bool | None):
        if value is None:
            self._remove(self._props.grid)
            return
        self._set(self._props.grid, value)

    @property
    def prop_comment(self) -> bool | None:
        """
        Gets/Sets Print Comments.
        """
        return self._get(self._props.comments)

    @prop_comment.setter
    def prop_comment(self, value: bool | None):
        if value is None:
            self._remove(self._props.comments)
            return
        self._set(self._props.comments, value)

    @property
    def prop_obj_img(self) -> bool | None:
        """
        Gets/Sets Print Object/images.
        """
        return self._get(self._props.obj_img)

    @prop_obj_img.setter
    def prop_obj_img(self, value: bool | None):
        if value is None:
            self._remove(self._props.obj_img)
            return
        self._set(self._props.obj_img, value)

    @property
    def prop_chart(self) -> bool | None:
        """
        Gets/Sets Print Charts.
        """
        return self._get(self._props.charts)

    @prop_chart.setter
    def prop_chart(self, value: bool | None):
        if value is None:
            self._remove(self._props.charts)
            return
        self._set(self._props.charts, value)

    @property
    def prop_drawing(self) -> bool | None:
        """
        Gets/Sets Print Drawing.
        """
        return self._get(self._props.draw_obj)

    @prop_drawing.setter
    def prop_drawing(self, value: bool | None):
        if value is None:
            self._remove(self._props.draw_obj)
            return
        self._set(self._props.draw_obj, value)

    @property
    def prop_formula(self) -> bool | None:
        """
        Gets/Sets Print Formulas.
        """
        return self._get(self._props.formula)

    @prop_formula.setter
    def prop_formula(self, value: bool | None):
        if value is None:
            self._remove(self._props.formula)
            return
        self._set(self._props.formula, value)

    @property
    def prop_zero_value(self) -> bool | None:
        """
        Gets/Sets Print Zero Values.
        """
        return self._get(self._props.zero_val)

    @prop_zero_value.setter
    def prop_zero_value(self, value: bool | None):
        if value is None:
            self._remove(self._props.zero_val)
            return
        self._set(self._props.zero_val, value)

    @property
    def _props(self) -> PrintProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = PrintProps(
                headers="PrintHeaders",
                grid="PrintGrid",
                comments="PrintAnnotations",
                obj_img="PrintObjects",
                charts="PrintCharts",
                draw_obj="PrintDrawing",
                formula="PrintFormulas",
                zero_val="PrintZeroValues",
            )
        return self._props_internal_attributes

    # endregion Properties
