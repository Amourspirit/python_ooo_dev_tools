from dataclasses import dataclass
from ooodev.utils.data_type.window_title import WindowTitle


@dataclass(frozen=True)
class DialogTitle(WindowTitle):
    """Dialog Title Info"""

    class_name: str = "SALSUBFRAME"
