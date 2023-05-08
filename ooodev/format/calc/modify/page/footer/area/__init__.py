import uno
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.preset.preset_image import PresetImageKind as PresetImageKind
from ooodev.utils.data_type.angle import Angle as Angle
from ooodev.utils.data_type.color_range import ColorRange as ColorRange
from ooodev.utils.data_type.size_mm import SizeMM as SizeMM
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind as ImgStyleKind
from ooodev.format.inner.common.format_types.size_percent import SizePercent as SizePercent
from ooodev.format.inner.common.format_types.offset_row import OffsetRow as OffsetRow
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn as OffsetColumn
from ooodev.format.inner.modify.calc.page.footer.area.color import InnerColor as InnerColor
from ooodev.format.inner.modify.calc.page.footer.area.color import Color as Color
from ooodev.format.inner.modify.calc.page.footer.area.img import Img as Img

__all__ = ["Color", "Img"]
