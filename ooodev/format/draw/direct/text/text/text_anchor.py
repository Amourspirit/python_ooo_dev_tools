from __future__ import annotations
import uno

from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind
from ooodev.format.inner.direct.draw.shape.text.text.text_anchor import TextAnchor as ShapeTextAnchor


class TextAnchor(ShapeTextAnchor):
    """
    This class represents the text spacing of an object that supports ``com.sun.star.drawing.TextProperties``.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_text_text_anchor`

    .. versionadded:: 0.17.5
    """

    def __init__(
        self,
        anchor_point: ShapeBasePointKind | None = None,
        full_width: bool | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            anchor_point (ShapeBasePointKind, optional): Anchor Point.
            full_width (bool, optional): Full Width. Defaults to None.

        Returns:
            None:

        Note:
            ``full_width`` applies when ``anchor_point`` is ``None`` or
            ``ShapeBasePointKind.TOP_CENTER`` or ``ShapeBasePointKind.CENTER``
            or ``ShapeBasePointKind.BOTTOM_CENTER``.

        See Also:
            - :ref:`help_draw_format_direct_shape_text_text_anchor`
        """
        super().__init__(anchor_point=anchor_point, full_width=full_width)
