from enum import Enum


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
