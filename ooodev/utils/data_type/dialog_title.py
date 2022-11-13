from dataclasses import dataclass
from .window_title import WindowTitle


@dataclass(frozen=True)
class DialogTitle(WindowTitle):
    """Dialog Title Info"""

    class_name: str = "SALSUBFRAME"
