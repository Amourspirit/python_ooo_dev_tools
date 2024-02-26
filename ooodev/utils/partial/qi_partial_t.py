from __future__ import annotations
from typing import Type, Literal, overload, TYPE_CHECKING, Optional, TypeVar

T = TypeVar("T")

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class QiPartialT(Protocol):
    @overload
    def qi(self, atype: Type[T]) -> Optional[T]:  # pylint: disable=invalid-name
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type such as XInterface

        Returns:
            T | None: instance of interface if supported; Otherwise, None
        """
        ...

    @overload
    def qi(self, atype: Type[T], raise_err: Literal[True]) -> T:  # pylint: disable=invalid-name
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type such as XInterface
            raise_err (bool, optional): If True then raises MissingInterfaceError if result is None. Default False

        Raises:
            MissingInterfaceError: If 'raise_err' is 'True' and result is None

        Returns:
            T: instance of interface.
        """
        ...

    @overload
    def qi(self, atype: Type[T], raise_err: Literal[False]) -> Optional[T]:  # pylint: disable=invalid-name
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type such as XInterface
            raise_err (bool, optional): If True then raises MissingInterfaceError if result is None. Default False

        Raises:
            MissingInterfaceError: If 'raise_err' is 'True' and result is None

        Returns:
            T | None: instance of interface if supported; Otherwise, None
        """
        ...
