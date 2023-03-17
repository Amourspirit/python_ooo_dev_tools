# region imports
from __future__ import annotations

from ...style_base import StyleBase

# endregion imports


class StructBase(StyleBase):
    """
    Struct Base Class.
    """

    # region Overrides
    def _get_internal_cattribs(self) -> dict:
        ca = super()._get_internal_cattribs()
        ca["_property_name"] = self._get_property_name()
        return ca

    # endregion Overrides

    # region Internal Methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            raise NotImplementedError

    # endregion Internal Methods
