from __future__ import annotations
from typing import cast, TYPE_CHECKING
from enum import Enum
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.reflection.enum_type_description_partial import EnumTypeDescriptionPartial

if TYPE_CHECKING:
    from com.sun.star.reflection import XEnumTypeDescription


class EnumTypeDescriptionComp(ComponentProp, EnumTypeDescriptionPartial):
    """
    Class for managing XEnumTypeDescription.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEnumTypeDescription) -> None:
        """
        Constructor

        Args:
            component (XEnumTypeDescription): UNO Component that implements ``com.sun.star.reflection.XEnumTypeDescription`` interface.
        """
        ComponentProp.__init__(self, component)
        EnumTypeDescriptionPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by TypeDescriptionEnumerationPartial
        return ()

    # endregion Overrides

    def get_name_value_dict(self) -> dict[str, int]:
        """
        Returns the enum member names and values as a dictionary.
        """
        names = self.get_enum_names()
        values = self.get_enum_values()
        return dict(zip(names, values))

    def get_value_name_dict(self) -> dict[int, str]:
        """
        Returns the enum member values and names as a dictionary.
        """
        names = self.get_enum_names()
        values = self.get_enum_values()
        return dict(zip(values, names))

    def get_name_from_value(self, value: int) -> str:
        """
        Returns the enum member name from the value.

        Args:
            value (int): The enum member value.

        Returns:
            str: The enum member name, or an empty string if not found.
        """
        vn = self.get_value_name_dict()
        return vn[value] if value in vn else ""

    def create_dynamic_enum(self, name: str) -> Enum:
        """
        Returns a dynamic python enum from the current enum names and values.

        Args:
            name (str): The name of the dynamic enum.

        Returns:
            Enum: The dynamic python enum.

        Example:
            This example is when the enum is info is for ``com.sun.star.awt.FontSlant``.
            Any valid code name can be used for the dynamic enum.

            .. code-block:: python

                >>> my_enum = info.create_dynamic_enum("com.sun.star.awt.FontSlant")
                >>> print(my_enum)
                <enum 'com.sun.star.awt.FontSlant'>
                >>> for e in my_enum:
                ...    print(e.name, e.value)
                NONE NONE
                OBLIQUE OBLIQUE
                ITALIC ITALIC
                DONTKNOW DONTKNOW
                REVERSE_OBLIQUE REVERSE_OBLIQUE
                REVERSE_ITALIC REVERSE_ITALIC

        """
        names = self.get_enum_names()

        return Enum(name, dict(zip(names, names)))

    def create_dynamic_name_value_enum(self, name: str) -> Enum:
        """
        Returns a dynamic python enum from the current enum names and values.

        Args:
            name (str): The name of the dynamic enum.

        Returns:
            Enum: The dynamic python enum.

        Example:
            This example is when the enum is info is for ``com.sun.star.awt.FontSlant``.
            Any valid code name can be used for the dynamic enum.

            .. code-block:: python

                >>> my_enum = info.create_dynamic_enum("com.sun.star.awt.FontSlant")
                >>> print(my_enum)
                <enum 'com.sun.star.awt.FontSlant'>
                >>> for e in my_enum:
                ...    print(e.name, e.value)
                NONE 0
                OBLIQUE 1
                ITALIC 2
                DONTKNOW 3
                REVERSE_OBLIQUE 4
                REVERSE_ITALIC 5
        """
        return Enum(name, self.get_name_value_dict())

    # region Properties
    @property
    def component(self) -> XEnumTypeDescription:
        """XEnumTypeDescription Component"""
        # pylint: disable=no-member
        return cast("XEnumTypeDescription", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
