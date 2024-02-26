from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.adapter.container.name_access_t import NameAccessT
    from ooodev.adapter.container.index_container_t import IndexContainerT

    class NameIndexT(NameAccessT, IndexContainerT, Protocol):
        pass

else:
    NameIndexT = object
