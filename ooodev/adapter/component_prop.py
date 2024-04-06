from __future__ import annotations

from .component_base import ComponentBase


class ComponentProp(ComponentBase):

    @property
    def component(self):
        """Gets component for this instance."""
        # pylint: disable=no-member
        return self._ComponentBase__get_component()  # type: ignore
