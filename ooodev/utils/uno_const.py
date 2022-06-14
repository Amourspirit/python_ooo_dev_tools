# coding: utf-8
from __future__ import annotations
import os
from typing import Any, TYPE_CHECKING
import uno

_DOCS_BUILDING = os.environ.get("DOCS_BUILDING", None) == 'True'
# _DOCS_BUILDING is only true when sphinx is building docs.
# env var DOCS_BUILDING is set in docs/conf.py
_ON_RTD = os.environ.get('READTHEDOCS', None) == 'True'
# env var READTHEDOCS is true when read the docs is building
# maybe not needed as _DOCS_BUILDING is set in conf.py

if _DOCS_BUILDING or _ON_RTD:
    # mock class
    from sphinx.ext.autodoc.mock import _MockObject
    class UnoConst(_MockObject):
        pass
else:
    class UnoConst:
        """
        Get access to Uno Const values without having to directly import them.
        
        This is a singleton class.
        Repeated call to constructor for any given Enum type,
        such as ``com.sun.star.text.ControlCharacter`` will return the same object
        
        Example:
        
            .. code-block:: python
            
                ControlCharacter = UnoConst("com.sun.star.text.ControlCharacter")
                assert ControlCharacter.LINE_BREAK == 1
                
                MyConst = UnoConst("com.sun.star.text.ControlCharacter")
                assert ControlCharacter is MyConst # singleton, same instances

        See Also:
            :py:class:`~.uno_enum.UnoEnum`
        """
        _loaded = {}
        _initialized = False  # This class var is important. It is always False.
                            # The instances will override this with their own, 
                            # set to True.
        def __new__(cls, type_name: str):
            if type_name in cls._loaded:
                return cls._loaded[type_name]
            e = super().__new__(cls)
            e.__dict__["_type_name"] = type_name
            cls._loaded[type_name] = e
            return e
            
        def __init__(self, type_name: str) -> None:
            """
            Constructor
            
            Get access to Uno Const values without having to directly import them.

            Args:
                type_name (str): The namespace of the Const as a string. Such as 'com.sun.star.text.ControlCharacter'
            """
            self._initialized = True
        
        def __getattr__(self, __name: str) -> Any:
            if self._initialized:
                # Provide the caller attributes in whatever ways interest you.
                try:
                    key = __name
                    const = uno.getConstantByName(f"{self._type_name}.{__name}")
                    self.__dict__[key] = const
                    return self.__dict__[key]
                except Exception:
                    raise AttributeError(f"Const {self._type_name} has no attribute {__name}")
            else:
                try:
                    return self.__dict__[__name] # Transparent access to instance vars.
                except KeyError:
                    raise AttributeError(__name)

        def __setattr__(self, key, value):
            if self._initialized:
                pass # Provide caller ways to set attributes in whatever ways.
            else:
                self.__dict__[key] = value # Transparent access.

        @property
        def typeName(self) -> str:
            """
            Gets Uno Type Name

            Returns:
                str: Uno type Name
            """
            return self._type_name

    __all__ = ("UnoConst",)