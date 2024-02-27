from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_multi_t import StyleMultiT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from com.sun.star.chart2 import XChartDocument
    from ooodev.utils.color import Color
    from ooodev.units.angle import Angle as Angle
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_mm import UnitMM
    from ooo.dyn.drawing.hatch_style import HatchStyle
    from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
else:
    Protocol = object
    XChartDocument = Any
    Color = Any
    Angle = Any
    UnitT = Any
    UnitMM = Any
    HatchStyle = Any
    PresetHatchKind = Any


class ChartFillHatchT(StyleMultiT, Protocol):
    """Chart Fill Hatch Protocol"""

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: HatchStyle = ...,
        color: Color = ...,
        space: float | UnitT = ...,
        angle: Angle | int = ...,
        bg_color: Color = ...,
    ) -> None:
        """
        Constructor.

        Args:
            chart_doc (XChartDocument): Chart document.
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            space (float, UnitT, optional): Specifies the space between the lines in the hatch (in ``mm`` units) or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.
            bg_color(Color, optional): Specifies the background Color. Set this ``-1`` (default) for no background color.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetHatchKind) -> ChartFillHatchT: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetHatchKind, **kwargs) -> ChartFillHatchT: ...

    # region Properties
    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets the angle."""
        ...

    @prop_angle.setter
    def prop_angle(self, value: Angle | int) -> None:
        """Sets the angle."""
        ...

    @property
    def prop_bg_color(self) -> Color:
        """Gets/Sets the background color."""
        ...

    @prop_bg_color.setter
    def prop_bg_color(self, value: Color) -> None:
        """Sets the background color."""
        ...

    @property
    def prop_color(self) -> Color:
        """Gets/Sets the color."""
        ...

    @prop_color.setter
    def prop_color(self, value: Color) -> None:
        """Sets the color."""
        ...

    @property
    def prop_style(self) -> HatchStyle:
        """Gets/Sets the style."""
        ...

    @prop_style.setter
    def prop_style(self, value: HatchStyle) -> None:
        """Sets the style."""
        ...

    @property
    def prop_space(self) -> UnitMM:
        """Gets/Sets the space."""
        ...

    @property
    def prop_hatch_name(self) -> str:
        """Gets Hatch Name."""
        ...

    @prop_space.setter
    def prop_space(self, value: float | UnitT) -> None:
        """Sets the space."""
        ...

    # endregion Properties
