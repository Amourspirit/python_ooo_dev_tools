from __future__ import annotations
from typing import Any
from ooodev.adapter.frame.the_desktop_comp import TheDesktopComp
from ooodev.loader.comp.components import Components


class TheDesktop(TheDesktopComp):
    """
    Class for managing theDesktop singleton Class.
    """

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.frame.theDesktop`` service.
        """
        TheDesktopComp.__init__(self, component=component)
        self._components = None

    @property
    def components(self) -> Components:
        """Desktop Components"""
        if self._components is None:
            self._components = Components(component=self.get_components())
        return self._components
