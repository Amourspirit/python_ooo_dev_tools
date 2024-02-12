from __future__ import annotations
from typing import TYPE_CHECKING

from ..structs.locale_t import LocaleT

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class FontLangT(LocaleT, Protocol):
    pass
