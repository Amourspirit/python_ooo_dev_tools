from __future__ import annotations
from typing import overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
else:
    Protocol = object


class ProtectT(StyleT, Protocol):
    """Size Protocol"""

    def __init__(
        self,
        *,
        position: bool | None = ...,
        size: bool | None = ...,
    ) -> None:
        """
        Constructor

        Args:
            position (bool, optional): Specifies position protection.
            size (bool, optional): Specifies size protection.

        Returns:
            None:
        """
        ...

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> ProtectT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> ProtectT: ...

    # endregion from_obj()

    # region Properties

    @property
    def prop_size(self) -> bool | None:
        """Gets/Sets size"""
        ...

    @prop_size.setter
    def prop_size(self, value: bool | None) -> None: ...

    @property
    def prop_position(self) -> bool | None:
        """Gets/Sets position"""
        ...

    @prop_position.setter
    def prop_position(self, value: bool | None) -> None: ...

    # endregion Properties
