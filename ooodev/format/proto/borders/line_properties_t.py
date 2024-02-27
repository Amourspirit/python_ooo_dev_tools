from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.format.inner.preset.preset_border_line import BorderLineKind
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_mm import UnitMM
    from ooodev.utils.color import Color
    from ooodev.utils.data_type.intensity import Intensity
else:
    Protocol = object
    BorderLineKind = Any
    UnitT = Any
    UnitMM = Any
    Color = Any
    Intensity = Any


class LinePropertiesT(StyleT, Protocol):
    """Size Protocol"""

    def __init__(
        self,
        *,
        style: BorderLineKind = ...,
        color: Color = ...,
        width: float | UnitT = ...,
        transparency: int | Intensity = ...,
    ) -> None:
        """
        Constructor

        Args:
            style (BorderLineKind): Line style. Defaults to ``BorderLineKind.CONTINUOUS``.
            color (Color, optional): Line Color. Defaults to ``Color(0)``.
            width (float | UnitT, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int | Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.

        Returns:
            None:
        """
        ...

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> LinePropertiesT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> LinePropertiesT: ...

    # endregion from_obj()

    # region Properties

    @property
    def prop_color(self) -> Color:
        """Gets/Sets the color."""
        ...

    @prop_color.setter
    def prop_color(self, value: Color): ...

    @property
    def prop_width(self) -> UnitMM: ...

    @prop_width.setter
    def prop_width(self, value: float | UnitT):
        """Gets/Sets the width."""
        ...

    @property
    def prop_style(self) -> BorderLineKind:
        """Gets/Sets the style."""
        ...

    @prop_style.setter
    def prop_style(self, value: BorderLineKind):
        """Sets the style."""
        ...

    @property
    def prop_transparency(self) -> Intensity:
        """Gets/Sets the transparency."""
        ...

    @prop_transparency.setter
    def prop_transparency(self, value: int | Intensity) -> None: ...

    # endregion Properties
