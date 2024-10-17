import uno  # noqa # type: ignore
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation
from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooodev.format.inner.direct.structs.side import (
    Side as Side,
    LineSize as LineSize,
    BorderLineKind as BorderLineKind,
)
from .shadow import Shadow as InnerShadow  # noqa # type: ignore
from .padding import Padding as InnerPadding  # noqa # type: ignore
from .borders import Borders as InnerBorders  # noqa # type: ignore
