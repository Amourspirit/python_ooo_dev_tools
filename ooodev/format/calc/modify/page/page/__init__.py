import uno
from ooo.dyn.style.page_style_layout import PageStyleLayout as PageStyleLayout
from ooo.dyn.style.numbering_type import NumberingTypeEnum as NumberingTypeEnum

from ooodev.utils.data_type.size_mm import SizeMM as SizeMM
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.preset.preset_paper_format import PaperFormatKind as PaperFormatKind
from ooodev.format.inner.direct.calc.page.page.margins import Margins as InnerMargins
from ooodev.format.inner.modify.calc.page.page.margins import Margins as Margins
from ooodev.format.inner.direct.write.page.page.paper_format import PaperFormat as InnerPaperFormat
from ooodev.format.inner.modify.calc.page.page.paper_format import PaperFormat as PaperFormat
from ooodev.format.inner.direct.calc.page.page.layout_settings import LayoutSettings as InnerLayoutSettings
from ooodev.format.inner.modify.calc.page.page.layout_settings import LayoutSettings as LayoutSettings

__all__ = ["Margins", "PaperFormat", "LayoutSettings"]
