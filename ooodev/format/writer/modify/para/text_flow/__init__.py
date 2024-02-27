import uno
from ooo.dyn.style.break_type import BreakType as BreakType

from ooodev.format.inner.modify.write.para.text_flow.breaks import Breaks as Breaks
from ooodev.format.inner.direct.write.para.text_flow.breaks import Breaks as InnerBreaks
from ooodev.format.inner.modify.write.para.text_flow.flow_options import FlowOptions as FlowOptions
from ooodev.format.inner.direct.write.para.text_flow.flow_options import FlowOptions as InnerFlowOptions
from ooodev.format.inner.modify.write.para.text_flow.hyphenation import Hyphenation as Hyphenation
from ooodev.format.inner.direct.write.para.text_flow.hyphenation import Hyphenation as InnerHyphenation
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind

__all__ = ["Breaks", "FlowOptions", "Hyphenation"]
