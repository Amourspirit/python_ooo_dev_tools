import uno  # noqa # type: ignore
from ooo.dyn.style.break_type import BreakType as BreakType

from ooodev.format.inner.modify.write.para.text_flow.breaks import Breaks as Breaks
from ooodev.format.inner.direct.write.para.text_flow.breaks import Breaks as InnerBreaks  # noqa: F401
from ooodev.format.inner.modify.write.para.text_flow.flow_options import (
    FlowOptions as FlowOptions,
)
from ooodev.format.inner.direct.write.para.text_flow.flow_options import (
    FlowOptions as InnerFlowOptions,  # noqa: F401
)
from ooodev.format.inner.modify.write.para.text_flow.hyphenation import (
    Hyphenation as Hyphenation,
)
from ooodev.format.inner.direct.write.para.text_flow.hyphenation import (
    Hyphenation as InnerHyphenation,  # noqa: F401
)
from ooodev.format.writer.style.para.kind.style_para_kind import (
    StyleParaKind as StyleParaKind,
)

__all__ = ["Breaks", "FlowOptions", "Hyphenation"]
