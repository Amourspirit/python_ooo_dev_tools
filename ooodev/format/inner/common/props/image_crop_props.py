from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class ImageCropProps(NamedTuple):
    """Internal Properties"""

    crop_struct: str
    width: str
    height: str
    size: str
    actual_size: str
