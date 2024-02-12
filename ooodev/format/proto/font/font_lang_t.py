from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ..structs.locale_t import LocaleT

if TYPE_CHECKING or DOCS_BUILDING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class FontLangT(LocaleT, Protocol):
    pass
