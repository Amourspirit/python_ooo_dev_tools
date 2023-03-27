import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle

from ooodev.format.inner.modify.write.page.transparency.gradient import Gradient as Gradient
from ooodev.format.inner.modify.write.page.transparency.gradient import InnerGradient as InnerGradient
from ooodev.format.inner.modify.write.page.transparency.transparency import InnerTransparency as InnerTransparency
from ooodev.format.inner.modify.write.page.transparency.transparency import Transparency as Transparency
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from ooodev.utils.data_type.angle import Angle as Angle
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange as IntensityRange
from ooodev.utils.data_type.offset import Offset as Offset

__all__ = ["Gradient", "InnerGradient", "InnerTransparency", "Transparency"]
