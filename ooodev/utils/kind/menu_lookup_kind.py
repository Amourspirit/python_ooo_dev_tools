from __future__ import annotations
from typing import Dict
from enum import Enum
from ooodev.utils.kind import kind_helper


class MenuLookupKind(Enum):
    """
    Represents DataSequenceRole

    See Also:
        `DataSequenceRole API <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1chart2_1_1data.html#a340775895b509a7f80ba895767123429>`_
    """

    FILE = ".uno:PickList"
    PICK_LIST = ".uno:PickList"
    TOOLS = ".uno:ToolsMenu"
    HELP = ".uno:HelpMenu"
    WINDOW = ".uno:WindowList"
    EDIT = ".uno:EditMenu"
    VIEW = ".uno:ViewMenu"
    INSERT = ".uno:InsertMenu"
    FORMAT = ".uno:FormatMenu"
    STYLES = ".uno:FormatStylesMenu"
    FORMAT_STYLES = ".uno:FormatStylesMenu"
    SHEET = ".uno:SheetMenu"
    DATA = ".uno:DataMenu"
    TABLE = ".uno:TableMenu"
    FORMAT_FORM = ".uno:FormatFormMenu"
    PAGE = ".uno:PageMenu"
    SHAPE = ".uno:ShapeMenu"
    SLIDE = ".uno:SlideMenu"
    SLIDESHOW = ".uno:SlideShowMenu"

    def __str__(self) -> str:
        return str(self.value)

    @staticmethod
    def from_str(s: str) -> "MenuLookupKind":
        """
        Gets an ``MenuLookupKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``MenuLookupKind`` instance.

        Returns:
            MenuLookupKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, MenuLookupKind)

    @staticmethod
    def get_dict(lower_keys: bool = False) -> Dict[str, str]:
        """
        Get ``MenuLookupKind`` Enum as dictionary

        Returns:
            dict[str, str]: Enum as dictionary
        """
        if lower_keys:
            return {k.name.lower(): k.value for k in MenuLookupKind}
        return {k.name: k.value for k in MenuLookupKind}
