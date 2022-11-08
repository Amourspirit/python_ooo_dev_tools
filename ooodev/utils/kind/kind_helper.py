# coding: utf-8
from __future__ import annotations
from enum import Enum
from typing import Any
from ...utils import gen_util as mGenUtil


def enum_class_new(cls, value: Any) -> Enum:
    """
    New (__new__) method for dynamically created enum classes

    Args:
        value (object): Can be Enum, Enum.value, str

    Raises:
        ValueError: if unable to match enum instance

    Returns:
        [Enum]: Enum Instance

    Example:
        ..code-block:: python

            >>> e = DrawingHatchingKind("YELLOW_45_DEGREES")
            >>> print(e.value)
            Yellow 45 Degrees
            >>> e = DrawingHatchingKind(DrawingHatchingKind.YELLOW_45_DEGREES)
            >>> print(e.value)
            Yellow 45 Degrees
            >>> e = DrawingHatchingKind(DrawingHatchingKind.YELLOW_45_DEGREES.value)
            >>> print(e.value)
            Yellow 45 Degrees
    """
    if isinstance(value, str):
        if hasattr(cls, value):
            return getattr(cls, value)
    _type = type(value)
    if _type is cls:
        return value
    raise ValueError("%r is not a valid %s" % (value, cls.__name__))


def enum_from_string(s: str, ec: Enum) -> Enum:
    """
    Gets an enum instance from a string

    Args:
        s (str): Name of enum instance.
            ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
            ``hypen-case``, ``normal case``.
        ec (Enum): Enum to get Enum instance from

    Raises:
        ValueError: If ``s`` string is empty.
        AttributeError: If unable to get enum instance.

    Returns:
        Enum: Enum Instance
    """
    if not s:
        raise ValueError("from_str arg s cannot be an empty value")

    try:
        return getattr(ec, s.upper())
    except AttributeError:
        pass

    e_str = mGenUtil.Util.to_single_space(s).replace("-", "_").replace(" ", "_")
    if "_" in e_str:
        e_str = e_str.upper()
    else:
        # handle pascalCase and CamelCase words
        e_str = mGenUtil.Util.to_snake_case_upper(e_str)
    return getattr(ec, e_str)
