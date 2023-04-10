from __future__ import annotations
from typing import Any, Tuple, cast, overload
import uno
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.hatch_struct import HatchStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset import preset_hatch as mPreset
from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.units import UnitMM
from ooodev.units import UnitObj
from ooodev.units.unit_convert import UnitConvert
from ooodev.utils import lo as mLo
from ooodev.utils.color import Color
from ooodev.utils.data_type.angle import Angle as Angle
from ooodev.utils.data_type.intensity import Intensity as Intensity

from ...chart.area.hatch import Hatch as ChartHatch


class Hatch(ChartHatch):
    """
    Class for Chart Title Area Fill Hatch.

    .. versionadded:: 0.9.4
    """

    pass
