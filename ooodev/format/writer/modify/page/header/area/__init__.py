import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

from ......preset.preset_gradient import PresetGradientKind as PresetGradientKind
from ......preset.preset_image import PresetImageKind as PresetImageKind
from ......preset.preset_hatch import PresetHatchKind as PresetHatchKind
from .......utils.data_type.angle import Angle as Angle
from .......utils.data_type.color_range import ColorRange as ColorRange
from .......utils.data_type.intensity import Intensity as Intensity
from .......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......direct.fill.area.img import ImgStyleKind as ImgStyleKind
from ......direct.common.format_types.size_percent import SizePercent as SizePercent
from ......direct.common.format_types.size_mm import SizeMM as SizeMM
from ......direct.common.format_types.offset_row import OffsetRow as OffsetRow
from ......direct.common.format_types.offset_column import OffsetColumn as OffsetColumn
from ......modify.page.header.area.color import Color as Color
from ......modify.page.header.area.gradient import Gradient as Gradient
from ......modify.page.header.area.img import Img as Img
from ......modify.page.header.area.hatch import Hatch as Hatch
