from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.inst.lo_inst import LoInst
from ooodev.adapter.xml.dom import document_builder_comp


def test_comp(loader, fix_lo_inst) -> None:
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = document_builder_comp.DocumentBuilderComp.from_lo(lo_inst=lo_inst)
    assert inst is not None
