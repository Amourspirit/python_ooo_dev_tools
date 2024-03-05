from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.format.inner.style_factory import write_area_img_factory
from .write_table_fill_img_partial import WriteTableFillImgPartial

if TYPE_CHECKING:
    from ooodev.format.proto.write.fill.area.fill_img_t import FillImgT
else:
    FillImgT = Any


class WriteFillImgPartial(WriteTableFillImgPartial):
    """
    Partial class for Write Fill Image.
    """

    def style_area_image_get(self) -> FillImgT | None:
        """
        Gets the Area Area Image Style.

        Raises:
            CancelEventError: If the event ``before_style_area_img_get`` is cancelled and not handled.

        Returns:
            FillImgT | None: Area image style or ``None`` if cancelled.
        """
        # mangled name
        # pylint: disable=no-member
        styler = self._WriteFillImgPartial__styler  # type: ignore
        return styler.style_get(factory=write_area_img_factory)
