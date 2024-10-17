import uno  # noqa # type: ignore
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ooodev.format.inner.direct.structs.side import BorderLineKind as BorderLineKind
from ooodev.format.inner.direct.structs.side import LineSize as LineSize
from ooodev.format.inner.direct.structs.side import Side as Side
from ooodev.format.inner.modify.write.para.border.borders import Borders as Borders
from ooodev.format.inner.direct.write.para.border.borders import Borders as InnerBorders  # noqa # type: ignore
from ooodev.format.inner.direct.write.para.border.padding import Padding as InnerPadding  # noqa # type: ignore
from ooodev.format.inner.modify.write.para.border.padding import Padding as Padding
from ooodev.format.inner.direct.write.para.border.shadow import Shadow as InnerShadow  # noqa # type: ignore
from ooodev.format.inner.modify.write.para.border.shadow import Shadow as Shadow
from ooodev.format.inner.direct.write.para.border.sides import Sides as InnerSides  # noqa # type: ignore
from ooodev.format.inner.modify.write.para.border.sides import Sides as Sides
from ooodev.format.writer.style.para.kind.style_para_kind import (
    StyleParaKind as StyleParaKind,
)

__all__ = ["Padding", "Shadow", "Sides"]
