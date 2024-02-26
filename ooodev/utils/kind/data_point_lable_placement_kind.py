"""
This module is DEPRECATED in version 0.27.0 It is no longer recommended for use and may be removed in the future.

Use :ref:`ooodev.utils.kind.data_point_label_placement_kind` instead.
"""

from ooodev.utils.kind.data_point_label_placement_kind import (
    DataPointLabelPlacementKind as DataPointLabelPlacementKind,
)
import warnings

warnings.warn(
    "ooodev.utils.kind.data_point_lable_placement_kind module is deprecated. Use ooodev.utils.kind.data_point_label_placement_kind.",
    DeprecationWarning,
    stacklevel=2,
)
