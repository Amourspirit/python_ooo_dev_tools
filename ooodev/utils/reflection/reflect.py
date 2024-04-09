from __future__ import annotations
from typing import Any

from ooodev.loader import lo as mLo


class Reflect:
    """
    Class for managing reflection.

    .. versionadded:: 0.37.0
    """

    @staticmethod
    def get_const_info(const_name: str) -> Any:
        """
        Gets information on a specific fully qualified type

        Args:
            const_name (str): The name of the constant.

        Returns:
            Any: The constant value.
        """
        from ooodev.adapter.reflection.the_type_description_manager_comp import TheTypeDescriptionManagerComp

        ctx = mLo.Lo.get_context()
        dm = TheTypeDescriptionManagerComp(
            ctx.getValueByName("/singletons/com.sun.star.reflection.theTypeDescriptionManager")
        )
        if dm.has_by_hierarchical_name(const_name):
            return dm.get_by_hierarchical_name(const_name)
        return None
