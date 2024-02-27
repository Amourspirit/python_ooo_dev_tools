import uno
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust as ParagraphAdjust
from ooo.dyn.text.paragraph_vert_align import ParagraphVertAlignEnum as ParagraphVertAlignEnum
from ooo.dyn.text.writing_mode2 import WritingMode2Enum as WritingMode2Enum

from ooodev.format.inner.direct.write.para.align.alignment import LastLineKind as LastLineKind
from ooodev.format.inner.direct.write.para.align.writing_mode import WritingMode as WritingMode
from ooodev.format.inner.modify.write.para.align.alignment import Alignment as Alignment
from ooodev.format.inner.direct.write.para.align.alignment import Alignment as InnerAlignment
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind

__all__ = ["Alignment"]
