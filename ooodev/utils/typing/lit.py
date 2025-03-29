from __future__ import annotations
from typing import Any, TYPE_CHECKING

# Literal is # Py >= 3.8

if TYPE_CHECKING:
    from typing_extensions import Literal as Literal
else:
    Literal = Any
