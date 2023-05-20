from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class UnoInterface(Protocol):
    """
    Protocol Class for UNO Interfaces
    """

    __pyunointerface__: str


class UnoStruct(Protocol):
    """
    Protocol Class for UNO Structs
    """

    typeName: str


class UnoException(Protocol):
    """
    Protocol Class for UNO Structs
    """

    __pyunointerface__: str
    typeName: str


class UnoEnum(Protocol):
    """
    Protocol Class for UNO Enums
    """

    typeName: str
    value: Any

    def __repr__(self) -> str:
        ...

    def __eq__(self, that: Any) -> bool:
        ...

    def __ne__(self, other: Any) -> bool:
        ...


class UnoType(Protocol):
    """
    Protocol Class for a UNO type.
    """

    typeName: str
    typeClass: UnoEnum

    def __init__(self, typeName: str, typeClass: str) -> None:
        ...

    def __repr__(self) -> str:
        ...

    def __eq__(self, that: Any) -> bool:
        ...

    def __ne__(self, other: Any) -> bool:
        ...

    def __hash__(self) -> str:
        ...
