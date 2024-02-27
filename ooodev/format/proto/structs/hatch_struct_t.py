from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.utils.color import Color
    from ooodev.units.angle import Angle
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_mm import UnitMM
    from ooo.dyn.drawing.hatch import Hatch
    from ooo.dyn.drawing.hatch_style import HatchStyle
else:
    Protocol = object
    Color = Any
    Angle = Any
    UnitT = Any
    UnitMM = Any
    Hatch = Any
    HatchStyle = Any


class HatchStructT(StyleT, Protocol):
    """Hatch Struct Protocol"""

    def __init__(
        self,
        *,
        style: HatchStyle = ...,
        color: Color = ...,
        distance: float | UnitT = ...,
        angle: Angle | int = ...,
    ) -> None: ...

    """
        Constructor

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            distance (int, UnitT, optional): Specifies the distance between the lines in the hatch (in ``mm`` units) or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.

        Returns:
            None:
        """

    def get_uno_struct(self) -> Hatch:
        """
        Gets UNO ``Hatch`` from instance.

        Returns:
            Hatch: ``Hatch`` instance
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> HatchStructT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> HatchStructT: ...

    # region Style Properties

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_style(self) -> HatchStyle:
        """Gets/Sets the style of the hatch."""
        ...

    @prop_style.setter
    def prop_style(self, value: HatchStyle): ...

    @property
    def prop_color(self) -> Color:
        """Gets/Sets the color of the hatch lines."""
        ...

    @prop_color.setter
    def prop_color(self, value: Color): ...

    @property
    def prop_distance(self) -> UnitMM:
        """Gets/Sets the distance between the lines in the hatch (in ``mm`` units)."""
        ...

    @prop_distance.setter
    def prop_distance(self, value: float | UnitT): ...

    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets angle of the hatch."""
        ...

    @prop_angle.setter
    def prop_angle(self, value: Angle | int): ...

    # endregion Prop Properties
