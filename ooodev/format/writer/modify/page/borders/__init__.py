import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ooodev.format.inner.direct.structs.side import BorderLineKind as BorderLineKind
from ooodev.format.inner.direct.structs.side import LineSize as LineSize
from ooodev.format.inner.direct.structs.side import Side as Side
from ooodev.format.inner.direct.write.para.border.padding import Padding as InnerPadding
from ooodev.format.inner.modify.write.page.border.padding import Padding as Padding
from ooodev.format.inner.modify.write.page.border.shadow import InnerShadow as InnerShadow
from ooodev.format.inner.modify.write.page.border.shadow import Shadow as Shadow
from ooodev.format.inner.direct.write.para.border.sides import Sides as InnerSides
from ooodev.format.inner.modify.write.page.border.sides import Sides as Sides
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind

__all__ = ["Padding", "Shadow", "Sides"]
