import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ooodev.format.inner.direct.structs.side import BorderLineKind as BorderLineKind
from ooodev.format.inner.direct.structs.side import LineSize as LineSize
from ooodev.format.inner.direct.structs.side import Side as Side
from ooodev.format.inner.direct.write.para.border.padding import Padding as InnerPadding
from ooodev.format.inner.modify.write.frame.border.padding import Padding as Padding
from ooodev.format.inner.direct.write.para.border.shadow import Shadow as InnerShadow
from ooodev.format.inner.modify.write.frame.border.shadow import Shadow as Shadow
from ooodev.format.inner.direct.write.para.border.sides import Sides as InnerSides
from ooodev.format.inner.modify.write.frame.border.sides import Sides as Sides
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind as StyleFrameKind

__all__ = ["Padding", "Shadow", "Sides"]
