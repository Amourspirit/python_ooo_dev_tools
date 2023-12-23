"""
Module for managing shape paragraph alignment.

.. versionadded:: 0.17.8
"""
from __future__ import annotations
import uno
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust

from ooodev.format.inner.direct.write.para.align.alignment import LastLineKind
from ooodev.format.inner.direct.draw.shape.para.alignment.alignment import Alignment as ShapeAlignment


class Alignment(ShapeAlignment):
    """
    Shape Paragraph Alignment

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_paragraph_alignment`

    .. versionadded:: 0.17.8
    """

    # region init

    def __init__(
        self,
        *,
        align: ParagraphAdjust | None = None,
        align_last: LastLineKind | None = None,
        expand_single_word: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            align (ParagraphAdjust, optional): Determines horizontal alignment of a paragraph.
            align_last (LastLineKind, optional): Determines the adjustment of the last line.
            expand_single_word (bool, optional): Determines if single words are stretched.
                It is only valid if ``align`` and ``align_last`` are also valid.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_paragraph_alignment`
        """
        super().__init__(align=align, align_last=align_last, expand_single_word=expand_single_word)

    # endregion init
