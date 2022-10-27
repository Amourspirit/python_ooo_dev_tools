from enum import Enum


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
