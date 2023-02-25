import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

from ......preset.preset_gradient import PresetGradientKind as PresetGradientKind
from ......preset.preset_image import PresetImageKind as PresetImageKind
from ......preset.preset_hatch import PresetHatchKind as PresetHatchKind
from ......preset.preset_pattern import PresetPatternKind as PresetPatternKind
from .......utils.data_type.angle import Angle as Angle
from .......utils.data_type.color_range import ColorRange as ColorRange
from .......utils.data_type.intensity import Intensity as Intensity
from .......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......direct.fill.area.img import ImgStyleKind as ImgStyleKind
from ......direct.common.format_types.size_percent import SizePercent as SizePercent
from .......utils.data_type.size_mm import SizeMM as SizeMM
from ......direct.common.format_types.offset_row import OffsetRow as OffsetRow
from ......direct.common.format_types.offset_column import OffsetColumn as OffsetColumn
from ......modify.page.header.area.color import InnerColor as InnerColor
from ......modify.page.footer.area.color import Color as Color
from ......modify.page.header.area.gradient import InnerGradient as InnerGradient
from ......modify.page.footer.area.gradient import Gradient as Gradient
from ......modify.page.header.area.img import InnerImg as InnerImg
from ......modify.page.footer.area.img import Img as Img
from ......modify.page.header.area.hatch import InnerHatch as InnerHatch
from ......modify.page.footer.area.hatch import Hatch as Hatch
from ......modify.page.header.area.pattern import InnerPattern as InnerPattern
from ......modify.page.footer.area.pattern import Pattern as Pattern
