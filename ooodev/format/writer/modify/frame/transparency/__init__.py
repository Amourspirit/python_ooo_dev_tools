import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle

from ooodev.format.inner.modify.write.frame.transparent.gradient import Gradient as Gradient
from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as InnerTransparency
from ooodev.format.inner.modify.write.frame.transparent.transparency import Transparency as Transparency
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind as StyleFrameKind
from ooodev.units.angle import Angle as Angle
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange as IntensityRange
from ooodev.utils.data_type.offset import Offset as Offset

__all__ = ["Gradient", "Transparency"]
