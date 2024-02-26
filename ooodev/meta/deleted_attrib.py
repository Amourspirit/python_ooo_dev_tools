from __future__ import annotations

# https://stackoverflow.com/questions/23181442/how-to-hide-remove-some-methods-in-inherited-class-in-python
from ooodev.exceptions import ex as mEx


class DeletedAttrib:
    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, instance, owner):
        cls_name = owner.__name__
        accessed_via = f"type object {cls_name!r}" if instance is None else f"{cls_name!r} object"
        raise mEx.DeletedAttributeError(f"attribute {self.name!r} of {accessed_via} has been deleted")
