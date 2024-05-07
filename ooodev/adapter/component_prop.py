from __future__ import annotations

from ooodev.adapter.component_base import ComponentBase


class ComponentProp(ComponentBase):

    def __bool__(self) -> bool:
        return self.component is not None

    def __eq__(self, value: object) -> bool:
        if value is self:
            return True
        try:
            if not isinstance(value, ComponentProp):
                return NotImplemented
        except TypeError:
            return False
        return self.component == value.component

    @property
    def component(self):
        """Gets component for this instance."""
        # pylint: disable=no-member
        return self._ComponentBase__get_component()  # type: ignore
