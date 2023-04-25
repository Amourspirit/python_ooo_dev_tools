import uno
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ooodev.format.writer.style.char import StyleCharKind as StyleCharKind
from ooodev.format.inner.direct.structs.side import BorderLineKind as BorderLineKind
from ooodev.format.inner.direct.structs.side import LineSize as LineSize
from ooodev.format.inner.direct.structs.side import Side as Side
from ooodev.format.inner.modify.write.char.border.sides import InnerSides as InnerSides
from ooodev.format.inner.modify.write.char.border.sides import Sides as Sides
from ooodev.format.inner.modify.write.char.border.padding import InnerPadding as InnerPadding
from ooodev.format.inner.modify.write.char.border.padding import Padding as Padding
from ooodev.format.inner.modify.write.char.border.shadow import InnerShadow as InnerShadow
from ooodev.format.inner.modify.write.char.border.shadow import Shadow as Shadow

__all__ = ["Sides", "Padding", "Shadow"]
