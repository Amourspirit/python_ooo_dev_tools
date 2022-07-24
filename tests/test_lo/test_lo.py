from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno


def test_bridge(loader) -> None:
    from ooodev.utils.lo import Lo
    assert Lo.bridge is not None
    assert Lo.is_loaded
    # don't know how to pass doc string to custom classproperty yet
    doc_str = Lo.bridge.__doc__
    assert True
    

def _test_dispose() -> None:
    # this test is to be run manually
    # if this test were to be run with other test
    # it would wipe out the connection to office because Lo is a static class.
    from ooodev.utils.lo import Lo
    with Lo.Loader(Lo.ConnectPipe(headless=True)) as loader:
        # confirm that Lo is connected to office
        assert Lo._xcc is not None
        assert Lo._mc_factory is not None
        assert Lo._lo_inst is not None
    # delay to ensure office bridge is disposed
    Lo.delay(1000)
    # now that bridge is disposed Lo internals should be reset.
    # this is because when Lo.load_office is called a listener is attached to
    # bridge via LoNamedEvent.OFFICE_LOADED event.
    assert Lo._xcc is None
    assert Lo._mc_factory is None
    assert Lo._lo_inst is None