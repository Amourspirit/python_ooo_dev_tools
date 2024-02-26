from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.structs.locale_t import LocaleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
else:
    Protocol = object


class FontLangT(LocaleT, Protocol):
    pass
