from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.helper.dot_dict import DotDict


def test_init() -> None:
    d = DotDict[str](a="hello", b="world")
    assert d.a == "hello"
    assert d.b == "world"
    assert d["a"] == "hello"
    assert d["b"] == "world"


def test_mixed_types() -> None:
    d = DotDict[object](a="hello", b=42, c=3.14, d=True)
    assert d.a == "hello"
    assert d.b == 42
    assert d.c == 3.14

    assert d.d is True

    assert len(d) == 4


def test_mixed_no_generic_types() -> None:
    d = DotDict(a="hello", b=42, c=3.14, d=True)
    assert d.a == "hello"
    assert d.b == 42
    assert d.c == 3.14
    assert d.d is True


def test_missing_attribute() -> None:
    d = DotDict[str](a="hello")
    with pytest.raises(AttributeError):
        _ = d.missing

    # Test with missing_attr_val
    d = DotDict[str](missing_attr_val="default", a="hello")
    assert d.missing == "default"

    d = DotDict[str](None, a="hello")
    assert d.missing is None


def test_set_get_del_attribute() -> None:
    d = DotDict[str]()
    d.test = "value"
    assert d.test == "value"
    assert d["test"] == "value"
    assert len(d) == 1

    del d.test
    assert "test" not in d
    assert len(d) == 0
    with pytest.raises(AttributeError):
        _ = d.test


def test_set_get_del_item() -> None:
    d = DotDict[int]()
    d["test"] = 42
    assert d["test"] == 42
    assert d.test == 42

    del d["test"]
    assert "test" not in d
    with pytest.raises(KeyError):
        _ = d["test"]


def test_contains_len() -> None:
    d = DotDict[str](a="hello", b="world")
    assert "a" in d
    assert "c" not in d
    assert len(d) == 2


def test_copy() -> None:
    original = DotDict[str](a="hello", b="world")
    copied = original.copy()

    assert len(original) == 2
    assert len(copied) == 2

    assert original.a == copied.a
    assert original.b == copied.b

    for key in original.keys():
        assert key in copied

    for key in copied.keys():
        assert key in original

    # Modify copy shouldn't affect original
    copied.c = "new"
    assert "c" not in original

    assert len(original) == 2
    assert len(copied) == 3


def test_dict_methods() -> None:
    d = DotDict[str](a="hello", b="world")

    # Test get
    assert d.get("a") == "hello"
    assert d.get("missing") is None
    assert d.get("missing", "default") == "default"

    # Test items
    items = dict(d.items())
    assert items == {"a": "hello", "b": "world"}

    # Test keys
    assert set(d.keys()) == {"a", "b"}

    # Test values
    assert set(d.values()) == {"hello", "world"}


def test_update() -> None:
    d1 = DotDict[str](a="hello")
    d2 = DotDict[str](b="world")
    d1.update(d2)
    assert d1.a == "hello"
    assert d1.b == "world"

    assert len(d1) == 1
    assert len(d2) == 1

    # Test update with regular dict
    d1.update({"c": "!"})
    assert d1.c == "!"
    assert len(d1) == 2

    d3 = DotDict[str](a="nice", b="day")
    d3.c = "!"
    d2.update(d3)
    assert d2.a == "nice"
    assert d2.b == "day"
    assert d2.c == "!"
    assert len(d2) == 3


def test_clear() -> None:
    d = DotDict[str](a="hello", b="world")
    d.clear()
    assert len(d) == 0
    assert not list(d.keys())


def test_copy_dict() -> None:
    d = DotDict[str](a="hello", b="world")
    copied_dict = d.copy_dict()
    assert isinstance(copied_dict, dict)
    assert copied_dict == {"a": "hello", "b": "world"}


def test_keys() -> None:
    d = DotDict[str](a="hello")
    d.b = "world"
    for key in d.keys():
        assert key in ["a", "b"]


def test_override_keys() -> None:
    d = DotDict[str](a="hello", keys="world")
    assert d.keys == "world"


def test_override_copy() -> None:
    d = DotDict[str](a="hello", copy="world")
    assert d.copy == "world"


def test_override_items() -> None:
    d = DotDict[object](a="hello", items=["this", "that"])
    assert d.items == ["this", "that"]
