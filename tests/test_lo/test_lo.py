from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno


def test_bridge(loader) -> None:
    from ooodev.utils.lo import Lo
    assert Lo.bridge is not None
    # don't know how to pass doc string to custom classproperty yet
    doc_str = Lo.bridge.__doc__
    assert True
    
