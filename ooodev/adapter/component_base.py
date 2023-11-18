from __future__ import annotations
from typing import Any
from abc import ABC

import uno  # pylint: disable=unused-import
from com.sun.star.lang import XComponent
from ooodev.events.args.generic_args import GenericArgs


class ComponentBase(ABC):
    """
    Base Class for Components in the ``component`` name space.
    """

    def __init__(self, component: Any) -> None:
        self._set_component(component)

    def _set_component(self, component: XComponent) -> None:
        self._component = component

    def _get_component(self) -> XComponent:
        return self._component

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def __getattr__(self, name: str) -> Any:
        comp = self._get_component()
        if hasattr(comp, name):
            return getattr(comp, name)
        raise AttributeError(name)
