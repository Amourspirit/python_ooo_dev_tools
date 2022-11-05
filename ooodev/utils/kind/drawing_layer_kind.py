from enum import Enum
from . import kind_helper


class DrawingLayerKind(str, Enum):
    """
    Drawing Layer Values.

    See Also:
        :py:meth:`~.Draw.get_layer`
    """

    BACK_GROUND = "background"
    BACK_GROUND_OBJECTS = "backgroundobjects"
    CONTROLS = "controls"
    LAYOUT = "layout"
    MEASURE_LINES = "measurelines"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "DrawingLayerKind":
        """
        Gets an ``DrawingLayerKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hypen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingLayerKind`` instance.

        Returns:
            DrawingLayerKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DrawingLayerKind)
