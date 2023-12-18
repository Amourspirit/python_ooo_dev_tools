from ooodev.units.angle import Angle as Angle
import warnings

# Deprecation in 0.17.4

warnings.warn(
    "ooodev.utils.data_type.angle Angle class is deprecated. Use ooodev.units Angle instead.",
    DeprecationWarning,
    stacklevel=2,
)
