import uno
from ooo.dyn.style.page_style_layout import PageStyleLayout as PageStyleLayout
from ooo.dyn.style.numbering_type import NumberingTypeEnum as NumberingTypeEnum

from ooodev.format.inner.direct.write.page.page.layout_settings import LayoutSettings as InnerLayoutSettings
from ooodev.format.inner.modify.write.page.page.layout_settings import LayoutSettings as LayoutSettings
from ooodev.format.inner.direct.write.page.page.margins import Margins as InnerMargins
from ooodev.format.inner.modify.write.page.page.margins import Margins as Margins
from ooodev.format.inner.direct.write.page.page.paper_format import PaperFormat as InnerPaperFormat
from ooodev.format.inner.modify.write.page.page.paper_format import PaperFormat as PaperFormat
from ooodev.format.inner.preset.preset_paper_format import PaperFormatKind as PaperFormatKind
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ooodev.utils.data_type.size_mm import SizeMM as SizeMM

__all__ = ["LayoutSettings", "Margins", "PaperFormat"]
