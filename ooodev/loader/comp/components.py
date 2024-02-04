from __future__ import annotations
from typing import Any
from ooodev.adapter.frame.components_comp import ComponentsComp


class Components(ComponentsComp):
    """
    Class for managing Components.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.frame.Components`` service.
        """
        ComponentsComp.__init__(self, component=component)
