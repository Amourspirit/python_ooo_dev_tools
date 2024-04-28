from __future__ import annotations
from typing import Any, Tuple, Protocol
import uno


class TypeProviderT(Protocol):
    """
    Protocol class for ``XTypeProvider``.
    """

    # region XTypeProvider
    def getImplementationId(self) -> uno.ByteSequence:
        """
        Obsolete unique identifier.

        Originally returned a sequence of bytes which, when non-empty,
        was used as an ID to distinguish unambiguously between two sets of types,
        for example to realize hashing functionality when the object is introspected.
        Two objects that returned the same non-empty ID had to return the same set of types in getTypes().
        (If a unique ID could not be provided, this method was always allowed to return an empty sequence, though).
        """
        ...

    def getTypes(self) -> Tuple[Any, ...]:
        """
        Returns a sequence of all types (usually interface types) provided by the object.
        """
        ...

    # endregion XTypeProvider
