from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class PropPartialT(Protocol):
    """Type for PropPartial Class."""

    @overload
    def get_property(self, name: str) -> Any:
        """
        Get property value

        Args:
            name (str): Property Name.

        Returns:
            Any: Property value or default.
        """
        ...

    @overload
    def get_property(self, name: str, default: Any) -> Any:
        """
        Get property value

        Args:
            name (str): Property Name.
            default (Any, optional): Return value if property value is ``None``.

        Returns:
            Any: Property value or default.
        """
        ...

    def set_property(self, **kwargs: Any) -> None:
        """
        Set property value

        Args:
            **kwargs: Variable length Key value pairs used to set properties.
        """
        ...
