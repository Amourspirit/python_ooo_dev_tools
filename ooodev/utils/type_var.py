# coding: utf-8
from typing import TypeVar
from os import PathLike

PathOrStr = TypeVar("PathOrStr", str, PathLike)
"""Path like object or string"""

UnoInterface = TypeVar("UnoInterface")
"""Represents a uno interface class. Any uno Class that starts with X"""

T = TypeVar("T")