from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.format.inner.direct.write.shape.area.shadow import Shadow as ShapeShadow

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.utils.color import Color
    from ooodev.format.inner.direct.write.shape.area.shadow import ShadowLocationKind


class Shadow(ShapeShadow):
    """
    Draw Shape Shadow

    .. seealso::

        - :ref:`help_draw_format_direct_shape_shadow`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        use_shadow: bool | None = None,
        location: ShadowLocationKind | None = None,
        color: Color | None = None,
        distance: float | UnitT | None = None,
        blur: int | UnitT | None = None,
        transparency: int | Intensity | None = None,
    ) -> None:
        """
        Constructor

        Args:
            use_shadow (bool, optional): Specifies if shadow is used.
            location (ShadowLocationKind , optional): Specifies the shadow location.
            color (Color , optional): Specifies shadow color.
            distance (float, UnitT , optional): Specifies shadow distance in ``mm`` units or :ref:`proto_unit_obj`.
            blur (int, UnitT, optional): Specifies shadow blur in ``pt`` units or in ``mm`` units  or :ref:`proto_unit_obj`.
            transparency (int , optional): Specifies shadow transparency value from ``0`` to ``100``.


        Returns:
            None:

        See Also:

            - :ref:`help_draw_format_direct_shape_shadow`
        """
        super().__init__(
            use_shadow=use_shadow,
            location=location,
            color=color,
            distance=distance,
            blur=blur,
            transparency=transparency,
        )
