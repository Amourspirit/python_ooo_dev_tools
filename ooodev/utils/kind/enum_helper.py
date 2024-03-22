from __future__ import annotations
from typing import cast
import uno
from com.sun.star.reflection import XEnumTypeDescription
from ooodev.adapter.reflection.enum_type_description_comp import EnumTypeDescriptionComp
from ooodev.loader import lo as mLo
from ooodev.utils.reflection.reflect import Reflect


class EnumHelper:
    """Class for managing enumeration."""

    @staticmethod
    def get_enum_info(name: str) -> EnumTypeDescriptionComp | None:
        """
        Gets information on a specific enumeration.

        Args:
            name (str): The name of the enumeration such as ``com.sun.star.awt.FontSlant``.

        Returns:
            EnumTypeDescriptionComp | None: The enumeration information, or ``None`` if not found.
        """
        if odt := cast(XEnumTypeDescription, Reflect.get_const_info(name)):
            return EnumTypeDescriptionComp(odt) if mLo.Lo.is_uno_interfaces(odt, XEnumTypeDescription) else None
        else:
            return None

    @classmethod
    def get_enum_name_from_value(cls, full_name: str, value: int) -> str:
        """
        Returns the enumeration member name from the value.

        Args:
            full_name (str): The full name of the enumeration such as ``com.sun.star.awt.FontSlant``.
            value (int): The enumeration member value.

        Raises:
            ValueError: If the enumeration is not found.

        Returns:
            str: The enumeration member name such as ``ITALIC``.
        """
        if info := cls.get_enum_info(full_name):
            return info.get_name_from_value(value)
        else:
            raise ValueError(f"Enumeration {full_name} not found.")

    @classmethod
    def get_uno_enum_from_value(cls, full_name: str, value: int) -> uno.Enum:
        """
        Returns the UNO enumeration from the value.

        Args:
            full_name (str): The full name of the enumeration such as ``com.sun.star.awt.FontSlant``.
            value (int): The enumeration member value.

        Raises:
            ValueError: If the enumeration is not found.

        Returns:
            uno.Enum: The UNO enumeration.
        """
        if info := cls.get_enum_info(full_name):
            name = info.get_name_from_value(value)
            return uno.Enum(full_name, name)
        else:
            raise ValueError(f"Enumeration {full_name} not found.")

    @classmethod
    def get_uno_enum_from_name(cls, full_name: str, name: str) -> uno.Enum:
        """
        Returns the UNO enumeration from the value.

        Args:
            full_name (str): The full name of the enumeration such as ``com.sun.star.awt.FontSlant``.
            name (str): The enumeration member name such as ``ITALIC``.

        Raises:
            ValueError: If the enumeration is not found.

        Returns:
            uno.Enum: The UNO enumeration.
        """
        try:
            return uno.Enum(full_name, name)
        except Exception as e:
            raise ValueError(f"Enumeration {full_name} not found.") from e
