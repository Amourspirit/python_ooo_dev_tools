# region Import
from __future__ import annotations

from ooodev.format.inner.style_base import StyleBase

# endregion Import


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
        # sourcery skip: raise-from-previous-error
        try:
            return self._property_name  # type: ignore
        except AttributeError:
            raise NotImplementedError

    def _set_property_name(self, name: str) -> str:
        # sourcery skip: raise-from-previous-error
        self._property_name = name

    # endregion Internal Methods
