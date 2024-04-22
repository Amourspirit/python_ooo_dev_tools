from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.inst.lo_inst import LoInst
from ooodev.adapter.frame import the_ui_command_description_comp


class _Dummy:
    pass


def test_comp(loader, fix_lo_inst) -> None:
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = the_ui_command_description_comp.TheUICommandDescriptionComp.from_lo(lo_inst=lo_inst)
    assert inst is not None
    key = "com.sun.star.frame.theUICommandDescription"
    assert key in lo_inst.cache
    # test cached version
    inst = the_ui_command_description_comp.TheUICommandDescriptionComp.from_lo(lo_inst=lo_inst)
    assert inst is not None


def test_builder(loader, fix_lo_inst):
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = the_ui_command_description_comp.TheUICommandDescriptionComp.from_lo(lo_inst=lo_inst)
    builder = the_ui_command_description_comp.get_builder(inst.component)
    assert builder is not None
    clz = builder.build_class(name="Test", base_class=_Dummy)
    assert clz is not None
