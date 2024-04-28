import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.helper.dot_dict import DotDict


def test_dot_dict():
    d = DotDict(a=1, b=2)
    assert d.a == 1
    assert d["b"] == 2
    d["c"] = 3
    assert d.c == 3
    assert "a" in d
    del d["a"]
    assert "a" not in d
    with pytest.raises(AttributeError):
        d.a
    d.a = 1
    assert d.a == 1
    assert d.get("a") == 1
    assert d.get("z") is None
