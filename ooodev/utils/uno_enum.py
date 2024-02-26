# coding: utf-8
from __future__ import annotations
import os
from typing import Any
import uno

from ooodev.mock import mock_g

if mock_g.DOCS_BUILDING:
    # mock class
    from sphinx.ext.autodoc.mock import _MockObject

    class UnoEnum(_MockObject):  # type: ignore
        pass

    # alternative below that includes a little more information.
    class _UnoEnum:
        # Note the __sphinx_mock__ attrib
        __sphinx_mock__ = True
        _loaded = {}
        _initialized = False

        def __new__(cls, type_name: str):
            if type_name in cls._loaded:
                return cls._loaded[type_name]
            e = super().__new__(cls)
            e.__dict__["_type_name"] = type_name
            cls._loaded[type_name] = e
            return e

        def __init__(self, type_name: str) -> None:
            self._initialized = True

        def __getattr__(self, __name: str) -> uno.Enum | Any:
            if self._initialized:
                # Provide the caller attributes in whatever ways interest you.
                try:
                    key = __name
                    e = f"<Enum instance {self._type_name} ({__name})>"
                    self.__dict__[key] = e
                    return self.__dict__[key]
                except Exception as exc:
                    raise AttributeError(f"Enum {self._type_name} has no attribute {__name}") from exc
            else:
                try:
                    return self.__dict__[__name]  # Transparent access to instance vars.
                except KeyError as e:
                    raise AttributeError(__name) from e

        def __setattr__(self, key, value):
            if not self._initialized:
                self.__dict__[key] = value  # Transparent access.

        @property
        def typeName(self) -> str:
            """
            Gets Uno Type Name

            Returns:
                str: Uno type Name
            """
            return self._type_name  # type: ignore

else:

    class UnoEnum:
        """
        Get access to Uno Enum values without having to directly import them.

        This is a singleton class.
        Repeated call to constructor for any given Enum type,
        such as ``com.sun.star.sheet.FillMode`` will return the same object

        Example:

            .. code-block:: python

                FillMode = UnoEnum("com.sun.star.sheet.FillMode")
                assert FillMode.LINEAR.value == "LINEAR"

                MyEnum = UnoEnum("com.sun.star.sheet.FillMode")
                assert FillMode is MyEnum # singleton, same instances

        Note:
            Uno enums can not be imported directly in python.

            ``from com.sun.star.sheet import FillMode`` is not a valid import
            and results in an import error.

            ``from com.sun.star.sheet.FillMode import LINEAR, SIMPLE, GROWTH`` is valid import but cumbersome.

            Best of both worlds. using ``typing.TYPE_CHECKING``.
            ``typing.TYPE_CHECKING`` is always false at runtime. This means anything imported
            in a ``typing.TYPE_CHECKING`` block is ignored at runtime.

            .. code-block:: python

                from typing import cast, TYPE_CHECKING
                if TYPE_CHECKING:
                    # uno enums are not able to be imported at runtime
                    from com.sun.star.sheet.FillMode as UnoFillMode

                FillMode = cast("UnoFillMode", UnoEnum("com.sun.star.sheet.FillMode"))

            In the above example ``FillMode`` will have full typing (intellisense) at design time.
            At runtime cast simply returns the UnoEnum instance.

            Note that ``"UnoFillMode"`` is wrapped in quotes. This is necessary for typing reasons.
            Without wrapping in quotes python will look for the import at runtime
            which will not be available because is in a ``TYPE_CHECKING`` block.

        See Also:
            :py:class:`~.uno_const.UnoConst`
        """

        _loaded = {}
        _initialized = False  # This class var is important. It is always False.

        # The instances will override this with their own,
        # set to True.
        def __new__(cls, type_name: str):
            # if _DOCS_BUILDING:
            #     # provision for sphinx autodoc
            #     # sphinx autodoc errors if this class is called
            #     # such as when a class instance is assigned to a constant
            #     # see Write class
            #     return type_name
            if type_name in cls._loaded:
                return cls._loaded[type_name]
            e = super().__new__(cls)
            e.__dict__["_type_name"] = type_name
            cls._loaded[type_name] = e
            return e

        def __init__(self, type_name: str) -> None:
            """
            Constructor

            Get access to Uno Enum values without having to directly import them.

            Args:
                type_name (str): The namespace of the enum as a string.
            """
            self._initialized = True

        def __getattr__(self, __name: str) -> uno.Enum | Any:
            if self._initialized:
                # Provide the caller attributes in whatever ways interest you.
                try:
                    key = __name
                    e = uno.Enum(self._type_name, __name)  # type: ignore
                    self.__dict__[key] = e
                    return self.__dict__[key]
                except Exception as exc:
                    raise AttributeError(f"Enum {self._type_name} has no attribute {__name}") from exc
            else:
                try:
                    return self.__dict__[__name]  # Transparent access to instance vars.
                except KeyError as e:
                    raise AttributeError(__name) from e

        def __setattr__(self, key, value):
            if not self._initialized:
                self.__dict__[key] = value  # Transparent access.

        @property
        def typeName(self) -> str:
            """
            Gets Uno Type Name

            Returns:
                str: Uno type Name
            """
            return self._type_name  # type: ignore

    __all__ = ("UnoEnum",)
