from __future__ import annotations

from ooodev.format.inner.direct.draw.shape.position_size.protect import Protect as ShapeProtect


class Protect(ShapeProtect):
    """
    Shape Protection

    .. seealso::

        - :ref:`help_draw_format_direct_shape_position_size_position_size_protect`

    .. versionadded:: 0.17.3
    """

    def __init__(
        self,
        position: bool | None = None,
        size: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            size (bool, optional): Specifies size protection.
            position (bool, optional): Specifies position protection.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_position_size_position_size_protect`
        """
        super().__init__(position=position, size=size)
