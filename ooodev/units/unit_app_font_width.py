from __future__ import annotations
from typing import Any, TYPE_CHECKING
from dataclasses import dataclass

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.units._app_font.unit_app_font_base import UnitAppFontBase
from ooodev.utils.kind.point_size_kind import PointSizeKind

if TYPE_CHECKING:
    from typing_extensions import Self
else:
    Self = Any

# pylint: disable=useless-parent-delegation


@dataclass(unsafe_hash=True)
class UnitAppFontWidth(UnitAppFontBase):
    """
    Unit in ``AppFont`` units.

    Supports ``UnitT`` protocol.

    Warning:
        Although this class support ``UnitT`` protocol, ``get_unit_length()`` method returns ``UnitLength.INVALID``.

    Note:
        Unlike most other units in this module, this unit is not based on ``UnitLength``.
        This means that it does not have a valid ``UnitLength`` value and returns ``UnitLength.INVALID``.

        This unit require that the application font pixel ratio be set before it can be used.
        Which means office must be loaded before this unit can be used.

    See Also:

        :ref:`proto_unit_obj`
    """

    def _set_ratio(self) -> None:
        # Because most all other unit module do not need to access Lo, it is imported here.
        # This should allow other modules to be imported without needing Lo,
        # that is, if they don't call get_app_font() method.
        # pylint: disable=import-outside-toplevel
        # pylint: disable=unsubscriptable-object
        from ooodev.loader.lo import Lo

        object.__setattr__(self, "_ratio", Lo.app_font_pixel_ratio.width)

    def get_value_oth_unit(self) -> float:
        """
        Return the Y value of the unit.

        Returns:
            float: The Y value of the unit. This is the value of a ``UnitAppFontY`` unit.
        """
        # pylint: disable=import-outside-toplevel
        # pylint: disable=unsubscriptable-object
        # convert to pixels and the apply the Y ratio
        from ooodev.loader.lo import Lo

        ratio = Lo.app_font_pixel_ratio.height
        px = self.get_value_px()
        return px * ratio

    def get_app_font_kind(self) -> PointSizeKind:
        """
        Gets the kind of the unit.

        Returns:
            PointSizeKind: Returns ``PointSizeKind.WIDTH``
        """
        return PointSizeKind.WIDTH

    # region math and comparison Overrides

    # not sure why but dataclasses does not recognize all __dunder__ methods unless they are overrides.
    @override
    def __int__(self) -> int:
        return super().__int__()

    @override
    def __eq__(self, other: object) -> bool:
        return super().__eq__(other)

    @override
    def __add__(self, other: object) -> Self:
        return super().__add__(other)

    @override
    def __radd__(self, other: object) -> Self:
        return super().__radd__(other)

    @override
    def __sub__(self, other: object) -> Self:
        return super().__sub__(other)

    @override
    def __rsub__(self, other: object) -> Self:
        return super().__rsub__(other)

    @override
    def __mul__(self, other: object) -> Self:
        return super().__mul__(other)

    @override
    def __rmul__(self, other: int) -> Self:
        return super().__rmul__(other)

    @override
    def __truediv__(self, other: object) -> Self:
        return super().__truediv__(other)

    @override
    def __rtruediv__(self, other: object) -> Self:
        return super().__rtruediv__(other)

    @override
    def __abs__(self) -> float:
        return super().__abs__()

    @override
    def __lt__(self, other: object) -> bool:
        return super().__lt__(other)

    @override
    def __le__(self, other: object) -> bool:
        return super().__le__(other)

    @override
    def __gt__(self, other: object) -> bool:
        return super().__gt__(other)

    @override
    def __ge__(self, other: object) -> bool:
        return super().__ge__(other)

    # endregion math and comparison Overrides
