import pytest
import time

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.cache.tlru_cache import TLRUCache


# Happy Path Tests
@pytest.mark.parametrize(
    "key, value, capacity, seconds, expected",
    [
        pytest.param("key1", "value1", 2, 5, "value1", id="single_item"),
        pytest.param("key2", "value2", 10, 10, "value2", id="under_capacity"),
        pytest.param("key3", "value3", 1, 1, "value3", id="at_capacity"),
    ],
    ids=str,
)
def test_put_and_get(key, value, capacity, seconds, expected):
    # Arrange
    cache = TLRUCache(capacity, seconds)

    # Act
    cache.put(key, value)
    result = cache.get(key)

    # Assert
    assert result == expected, "The value retrieved should match the value stored."


# Edge Cases
@pytest.mark.parametrize(
    "key, value, capacity, seconds, sleep_time, expected",
    [
        pytest.param("key1", "value1", 2, 1, 2, None, id="expired_item"),
        pytest.param("key2", "value2", 1, 5, 0, "value2", id="capacity_limit_reached_new_item_overwrites_old"),
    ],
    ids=str,
)
def test_time_expiry_and_capacity_limit(key, value, capacity, seconds, sleep_time, expected):
    # Arrange
    cache = TLRUCache(capacity, seconds)

    # Act
    cache.put(key, value)
    time.sleep(sleep_time)  # Wait for the item to potentially expire
    result = cache.get(key)

    # Assert
    assert result == expected, "The value should be None if expired or match the stored value."


# Error Cases
@pytest.mark.parametrize(
    "key, value, capacity, seconds, error",
    [
        pytest.param(None, "value1", 2, 5, TypeError, id="none_key"),
        pytest.param("key1", None, 2, 5, TypeError, id="none_value"),
    ],
    ids=str,
)
def test_errors(key, value, capacity, seconds, error):
    # Arrange
    if capacity > 0 and seconds > 0:
        cache = TLRUCache(capacity, seconds)
    else:
        with pytest.raises(error):
            TLRUCache(capacity, seconds)
        return

    # Act / Assert
    with pytest.raises(error):
        if key is None or value is None:
            cache.put(key, value)
        else:
            cache[key] = value


# Test Clear Method
@pytest.mark.parametrize(
    "capacity, seconds, items, expected_len_after_clear",
    [
        pytest.param(5, 5, [("key1", "value1"), ("key2", "value2")], 0, id="multiple_items"),
    ],
    ids=str,
)
def test_clear(capacity, seconds, items, expected_len_after_clear):
    # Arrange
    cache = TLRUCache(capacity, seconds)
    for key, value in items:
        cache.put(key, value)

    # Act
    cache.clear()

    # Assert
    assert len(cache) == expected_len_after_clear, "Cache should be empty after clear."


# Test Remove Method
@pytest.mark.parametrize(
    "key_to_remove, initial_items, expected_items",
    [
        pytest.param("key1", [("key1", "value1"), ("key2", "value2")], [("key2", "value2")], id="remove_existing"),
        pytest.param(
            "key3",
            [("key1", "value1"), ("key2", "value2")],
            [("key1", "value1"), ("key2", "value2")],
            id="remove_non_existing",
        ),
    ],
    ids=str,
)
def test_remove(key_to_remove, initial_items, expected_items):
    # Arrange
    cache = TLRUCache(len(initial_items), 10)
    for key, value in initial_items:
        cache.put(key, value)

    # Act
    cache.remove(key_to_remove)

    # Assert
    for key, _ in expected_items:
        assert cache.get(key) is not None, "Expected item was not found after removal."
    assert len(cache) == len(expected_items), "Cache size does not match expected after removal."
