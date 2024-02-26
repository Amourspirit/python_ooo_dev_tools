from __future__ import annotations
from typing import Any, TYPE_CHECKING, Union
from ooodev.loader.inst.lo_inst import LoInst

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class ComponentStructT(Protocol):
    component: Any


class ComponentInstStructT(ComponentStructT, Protocol):
    lo_inst: LoInst


class ComponentClassT(Protocol):
    @property
    def component(self) -> Any:
        raise NotImplementedError()


class ComponentInstClassT(ComponentClassT, Protocol):
    @property
    def lo_inst(self) -> LoInst:
        raise NotImplementedError()


if TYPE_CHECKING:
    ComponentT = Union[ComponentStructT, ComponentClassT]
    ComponentInstT = Union[ComponentInstStructT, ComponentInstClassT]
else:
    ComponentT = Any
    ComponentInstT = Any
