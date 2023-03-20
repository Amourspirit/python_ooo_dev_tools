import uno
from ooo.dyn.style.page_style_layout import PageStyleLayout as PageStyleLayout
from ooo.dyn.style.numbering_type import NumberingTypeEnum as NumberingTypeEnum
from ....style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ....style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from .....preset.preset_paper_format import PaperFormatKind as PaperFormatKind
from ......utils.data_type.size_mm import SizeMM as SizeMM
from .....modify.page.page.margins import Margins as Margins, InnerMargins as InnerMargins
from .....modify.page.page.paper_format import PaperFormat as PaperFormat, InnerPaperFormat as InnerPaperFormat
from .....modify.page.page.layout_settings import (
    LayoutSettings as LayoutSettings,
    InnerLayoutSettings as InnerLayoutSettings,
)
