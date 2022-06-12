# coding: utf-8
from typing import TypeVar, Union, TypeAlias
from os import PathLike

PathOrStr: TypeAlias = Union[str, PathLike]
"""Path like object or string"""

UnoInterface: TypeAlias = object
"""Represents a uno interface class. Any uno Class that starts with X"""

T = TypeVar("T")