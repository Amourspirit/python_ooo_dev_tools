from __future__ import annotations
from typing import cast, TypeVar

from ooodev.format.inner.common.props.page_margin_props import PageMarginProps
from ooodev.format.inner.direct.calc.page.page.margins import Margins as CalcMargins
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_obj import UnitT

_TMargins = TypeVar("_TMargins", bound="Margins")


class Margins(CalcMargins):
    """
    Page Margins.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
        gutter: float | UnitT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Left Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            right (float, optional): Right Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            top (float, optional): Top Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            bottom (float, optional): Bottom Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            gutter (float, optional): Gutter Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.

        Returns:
            None:
        """

        super().__init__(left=left, right=right, top=top, bottom=bottom)
        self.prop_gutter = gutter

    @property
    def prop_gutter(self) -> UnitMM | None:
        """Gets/Sets Gutter value"""
        pv = cast(int, self._get(self._props.gutter))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_gutter.setter
    def prop_gutter(self, value: float | UnitT | None) -> None:
        self._check(value, "gutter")
        if value is None:
            self._remove(self._props.gutter)
            return
        try:
            self._set(self._props.gutter, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.gutter, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def _props(self) -> PageMarginProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = PageMarginProps(
                left="LeftMargin", right="RightMargin", top="TopMargin", bottom="BottomMargin", gutter="GutterMargin"
            )
        return self._props_internal_attributes

    @property
    def default(self: _TMargins) -> _TMargins:  # type: ignore[misc]
        """Gets Margin Default. Static Property."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._DEFAULT_INST
        except AttributeError:
            inst = self.__class__(_cattribs=self._get_internal_cattribs())  # type: ignore
            for attrib in inst._props:
                if attrib:
                    inst._set(attrib, 2000)
            if inst._props.gutter:
                inst._set(inst._props.gutter, 0)
            inst._is_default_inst = True
            self._DEFAULT_INST = inst
        return self._DEFAULT_INST
