# region Imports
from __future__ import annotations
import uno

from ooodev.utils import props as mProps
from ooodev.exceptions import ex as mEx
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.modify.calc.cell_style_base import CellStyleBase
from ooodev.format.inner.modify.calc.page.sheet.scale_reduce_enlarge import ScaleProps

# endregion Imports


class ScaleNumOfPages(CellStyleBase):
    """
    Page Style Shrink Page Range.

    .. seealso::

        - :ref:`help_calc_format_modify_page_sheet`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        pages: int = 1,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            pages (int): Specifies the number of pages the spreadsheet will scale to when printed.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_sheet`
        """

        super().__init__(style_name=style_name, style_family=style_family)
        self.prop_pages = pages

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> ScaleNumOfPages:
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
            ScaleNumOfPages: ``ScaleNumOfPages`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        style_props = inst.get_style_props(doc)
        if not inst._is_valid_obj(style_props):
            raise mEx.NotSupportedError(f"Object is not support to convert to {cls.__name__}")

        inst.prop_pages = int(mProps.Props.get(style_props, inst._props.scale))
        return inst

    # region Properties
    @property
    def prop_pages(self) -> int:
        """
        Gets/Sets the number of pages the spreadsheet will scale to when printed.
        """
        return self._get(self._props.scale)

    @prop_pages.setter
    def prop_pages(self, value: int):
        # 1 is min
        value = max(value, 1)
        # the order is important here.
        # by setting value last it ensure it is the last property set.
        # Otherwise, may get unexpected results. This is a Calc issue.
        self._set(self._props.factor, 0)
        self._set(self._props.page_y, 0)
        self._set(self._props.page_x, 0)
        self._set(self._props.scale, value)

    @property
    def _props(self) -> ScaleProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ScaleProps(
                factor="PageScale", scale="ScaleToPages", page_x="ScaleToPagesX", page_y="ScaleToPagesY"
            )
        return self._props_internal_attributes

    # endregion Properties
