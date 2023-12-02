import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle

from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct as GradientStruct
from ooodev.utils.data_type.angle import Angle as Angle
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange as IntensityRange
from ooodev.utils.data_type.offset import Offset as Offset
from .gradient import Gradient as Gradient
from .transparency import Transparency as Transparency

__all__ = ["Gradient", "Transparency"]
