# coding: utf-8
"""
Uno Enum Helper Methods

These methods are used to add custom functionality to enums or enum like classes.
"""
import uno

def uno_enum_class_new(cls, value):
    """
    New (__new__) method for dynamically created uno.enum classes

    Args:
        value (object): Can be Uno.Enum, uno.Enum.value, str

    Raises:
        ValueError: if unable to match enum instance

    Returns:
        [uno.Enum]: Enum Instance

    Example:
        .. code-block:: python
        
            >>> e = HorizontalAlignment("RIGHT")
            >>> print(e.value)
            RIGHT
            >>> e = HorizontalAlignment(HorizontalAlignment.LEFT)
            >>> print(e.value)
            LEFT
            >>> e = HorizontalAlignment(HorizontalAlignment.CENTER.value)
            >>> print(e.value)
            CENTER
    """
    if isinstance(value, str):
        if hasattr(cls, value):
            return getattr(cls, value)
        # try for upper case match.
        uv = value.upper()
        if hasattr(cls, uv):
            return getattr(cls, uv)
    _type = type(value)
    if _type is uno.Enum:
        return value
    if _type is cls:
        return value
    raise ValueError("%r is not a valid %s" % (value, cls.__name__))


def uno_enum_class_ne(self, other: object) -> bool:
    """
    Enum Not equal method.
    
    This method is usuall assigned to an enum.

    Args:
        other (object): Object to compare

    Returns:
        bool: False if equal; Othherwise, true
    """
    return not self.__eq__(other)
