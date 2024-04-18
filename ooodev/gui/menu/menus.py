from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.gui.menu.menu_app import MenuApp
from ooodev.loader.inst.service import Service
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial


if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class Menus(LoInstPropsPartial):
    """Class for manager menus"""

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)

    def __getitem__(self, app: str | Service):
        """Index access"""
        return MenuApp(app=app, lo_inst=self.lo_inst)
