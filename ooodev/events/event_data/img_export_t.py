from __future__ import annotations
from typing import TypedDict
from ooodev.utils.type_var import PathOrStr


class ImgExportT(TypedDict):
    """
    Image Export Data

    .. versionadded:: 0.20.3
    """

    image_type: str
    compression: int
    interlaced: int
    translucent: int
    pixel_width: int
    pixel_height: int
    logical_width: int
    logical_height: int
    file: PathOrStr
