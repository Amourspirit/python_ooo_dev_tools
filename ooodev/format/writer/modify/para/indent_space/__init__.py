import uno
from ooodev.format.inner.direct.structs.line_spacing_struct import ModeKind as ModeKind
from ooodev.format.inner.modify.write.para.indent_space.indent import Indent as Indent
from ooodev.format.inner.direct.write.para.indent_space.indent import Indent as InnerIndent
from ooodev.format.inner.direct.write.para.indent_space.line_spacing import LineSpacing as InnerLineSpacing
from ooodev.format.inner.modify.write.para.indent_space.line_spacing import LineSpacing as LineSpacing
from ooodev.format.inner.direct.write.para.indent_space.spacing import Spacing as InnerSpacing
from ooodev.format.inner.modify.write.para.indent_space.spacing import Spacing as Spacing
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind

__all__ = ["Indent", "LineSpacing", "Spacing"]
