from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.uno_control_group_box_comp import UnoControlGroupBoxComp

if TYPE_CHECKING:
    from com.sun.star.form.control import GroupBox


class GroupBoxComp(UnoControlGroupBoxComp):
    """Class for GroupBox Control"""

    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.GroupBox",)

    @property
    @override
    def component(self) -> GroupBox:
        """GroupBox Component"""
        # pylint: disable=no-member
        return cast("GroupBox", self._ComponentBase__get_component())  # type: ignore
