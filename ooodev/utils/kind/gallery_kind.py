from enum import Enum
from . import kind_helper


class GalleryKind(str, Enum):
    """
    Gallery Kind.

    Used to encapsulate Gallery names.

    Note:
        Not all gallery name are necessarily lister in this enum.
        Different system and version may have a some different values.

    See Also:
        - :py:meth:`.Gallery.find_gallery_item`
        - :py:meth:`.Gallery.report_galeries`
    """

    BPMN = "BPMN"
    FLOW_CHART = "Flow chart"
    SOUNDS = "Sounds"
    ARROWS = "Arrows"
    DIAGRAMS = "Diagrams"
    NETWORK = "Network"
    SHAPES = "Shapes"
    BULLETS = "Bullets"
    ICONS = "Icons"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "GalleryKind":
        """
        Gets an ``GalleryKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hypen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``GalleryKind`` instance.

        Returns:
            GalleryKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, GalleryKind)
