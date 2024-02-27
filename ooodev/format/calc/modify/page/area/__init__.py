import uno
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

from ooodev.utils.data_type.offset import Offset as Offset
from ooodev.utils.data_type.size_mm import SizeMM as SizeMM
from ooodev.format.inner.common.format_types.size_percent import SizePercent as SizePercent
from ooodev.format.inner.common.format_types.offset_row import OffsetRow as OffsetRow
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn as OffsetColumn
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind as ImgStyleKind
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.preset.preset_image import PresetImageKind as PresetImageKind
from ooodev.format.inner.modify.calc.page.area.color import InnerColor as InnerColor
from ooodev.format.inner.modify.calc.page.area.color import Color as Color
from ooodev.format.inner.modify.calc.page.area.img import InnerImg as InnerImg
from ooodev.format.inner.modify.calc.page.area.img import Img as Img

__all__ = ["Color", "Img"]
