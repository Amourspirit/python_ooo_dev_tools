# region Imports
from __future__ import annotations
from typing import NamedTuple
import uno

from ooodev.utils import props as mProps
from ooodev.exceptions import ex as mEx
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.modify.calc.cell_style_base import CellStyleBase

# endregion Imports


class ScaleProps(NamedTuple):
    factor: str
    scale: str
    page_x: str
    page_y: str


class ScaleReduceEnlarge(CellStyleBase):
    """
    Page Style Scale Reduce or Enlarge.

    .. seealso::

        - :ref:`help_calc_format_modify_page_sheet`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        factor: int = 100,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            factor (int): Specifies scale factor between ``10`` and ``400``. Default is ``100``.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:
        See Also:
            - :ref:`help_calc_format_modify_page_sheet`
        """

        super().__init__(style_name=style_name, style_family=style_family)
        self.prop_factor = factor

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> ScaleReduceEnlarge:
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
            Order: ``Order`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        style_props = inst.get_style_props(doc)
        if not inst._is_valid_obj(style_props):
            raise mEx.NotSupportedError(f"Object is not support to convert to {cls.__name__}")

        inst.prop_factor = int(mProps.Props.get(style_props, inst._props.factor))
        return inst

    # region Properties
    @property
    def prop_factor(self) -> int:
        """
        Gets/Sets Scaling factor. Value from ``10`` to ``400``.
        """
        return self._get(self._props.factor)

    @prop_factor.setter
    def prop_factor(self, value: int):
        # 10% is min
        value = max(value, 10)
        # 400% is max
        value = min(value, 400)
        # the order is important here.
        # if factor is not last Calc will set it to 100
        # by setting it last here the property gets set last by Props
        self._set(self._props.scale, 0)
        self._set(self._props.page_x, 0)
        self._set(self._props.page_y, 0)
        self._set(self._props.factor, value)

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
