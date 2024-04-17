from __future__ import annotations

from .component_base import ComponentBase


class ComponentProp(ComponentBase):

    def __bool__(self) -> bool:
        return self.component is not None

    @property
    def component(self):
        """Gets component for this instance."""
        # pylint: disable=no-member
        return self._ComponentBase__get_component()  # type: ignore
