from __future__ import annotations
from ooodev.utils import props as mProps


class MenuDebug:
    """Class for debug info menu"""

    @classmethod
    def _get_info(cls, menu, index):
        """Get every option menu"""
        line = f"({index}) {menu.get('CommandURL', '----------')}"
        submenu = menu.get("ItemDescriptorContainer", None)
        if submenu is not None:
            line += cls._get_submenus(submenu)
        return line

    @classmethod
    def _get_submenus(cls, menu, level=1):
        """Get submenus"""
        line = ""
        for i, v in enumerate(menu):
            data = mProps.Props.data_to_dict(v)
            cmd = data.get("CommandURL", "----------")
            line += f'\n{"  " * level}├─ ({i}) {cmd}'
            submenu = data.get("ItemDescriptorContainer", None)
            if submenu is not None:
                line += cls._get_submenus(submenu, level + 1)
        return line

    def __call__(cls, menu):
        for i, m in enumerate(menu):
            data = mProps.Props.data_to_dict(m)
            print(cls._get_info(data, i))
        return
