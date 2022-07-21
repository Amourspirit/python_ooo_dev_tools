from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno


def test_bridge(loader) -> None:
    from ooodev.utils.lo import Lo
    assert Lo.bridge_component is not None
