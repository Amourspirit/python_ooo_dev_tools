from __future__ import annotations
from typing import Any, Type, Literal, overload, Optional, TypeVar
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.loader import lo as mLo

T = TypeVar("T")


class QiPartial:
    def __init__(self, component: Any, lo_inst: LoInst | None = None):
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__component = component

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

    # pylint: disable=invalid-name
    def qi(self, atype: Type[T], raise_err: bool = False) -> Optional[T]:
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type to query obj for. Any Uno class that starts with 'X' such as XInterface
            raise_err (bool, optional): If True then raises MissingInterfaceError if result is None. Default False

        Raises:
            MissingInterfaceError: If 'raise_err' is 'True' and result is None

        Returns:
            T | None: instance of interface if supported; Otherwise, None

        Note:
            When ``raise_err=True`` return value will never be ``None``.
        """
        if raise_err:
            return self.__lo_inst.qi(atype, self.__component, raise_err)
        return self.__lo_inst.qi(atype, self.__component)
