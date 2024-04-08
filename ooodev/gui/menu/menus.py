from __future__ import annotations
from ooodev.gui.menu.menu_app import MenuApp
from ooodev.loader.inst.service import Service


class Menus:
    """Class for manager menus"""

    def __getitem__(self, app: str | Service):
        """Index access"""
        return MenuApp(app)
