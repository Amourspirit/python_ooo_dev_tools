import uno  # noqa # type: ignore
from ooo.dyn.text.wrap_text_mode import WrapTextMode as WrapTextMode

from ooodev.format.inner.direct.write.frame.wrap.options import Options as InnerOptions  # noqa # type: ignore
from ooodev.format.inner.modify.write.frame.wrap.options import Options as Options
from ooodev.format.inner.direct.write.frame.wrap.settings import (
    Settings as InnerSettings,  # noqa # type: ignore
)
from ooodev.format.inner.modify.write.frame.wrap.settings import Settings as Settings
from ooodev.format.inner.direct.write.frame.wrap.spacing import Spacing as InnerSpacing  # noqa # type: ignore
from ooodev.format.inner.modify.write.frame.wrap.spacing import Spacing as Spacing
from ooodev.format.writer.style.frame.style_frame_kind import (
    StyleFrameKind as StyleFrameKind,
)

__all__ = ["Options", "Settings", "Spacing"]
