import pytest
from datetime import datetime, timedelta, timezone
import time

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.cache.time_cache import TimeCache


# Test __init__ method
@pytest.mark.parametrize(
    "seconds, expected_delta",
    [
        (60, 60),  # ID: Init-Seconds-Normal
        (0.1, 0.1),  # ID: Init-Seconds-Fraction
        (86400, 86400),  # ID: Init-Seconds-Day
    ],
)
def test_time_cache_init(seconds, expected_delta):
    # Arrange

    # Act
    cache = TimeCache(seconds=seconds)

    # Assert
    assert cache.seconds == expected_delta, "The delta seconds do not match the expected value."


# Test put and get methods
@pytest.mark.parametrize(
    "key, value, sleep_time, expected_result",
    [
        ("key1", "value1", 0, "value1"),  # ID: PutGet-NoDelay
        ("key2", "value2", 2, None),  # ID: PutGet-Expired
    ],
    ids=["NoDelay", "Expired"],
)
def test_put_get(key, value, sleep_time, expected_result):
    cache = TimeCache(seconds=1)

    # Arrange

    # Act
    cache.put(key, value)
    if sleep_time:
        time.sleep(sleep_time)
    result = cache.get(key)

    # Assert
    assert result == expected_result, "Cache get did not return the expected result."


# Test remove method
@pytest.mark.parametrize(
    "key, value, remove_key, expected_result",
    [
        ("key1", "value1", "key1", False),  # ID: Remove-Existing
        ("key2", "value2", "key3", True),  # ID: Remove-NonExisting
    ],
)
def test_remove(key, value, remove_key, expected_result):
    cache = TimeCache(seconds=10)

    # Arrange
    cache.put(key, value)

    # Act
    cache.remove(remove_key)
    result = key in cache

    # Assert
    assert result == expected_result, "Cache remove did not yield the expected result."


# Test clear method
def test_clear():
    cache = TimeCache(seconds=10)

    # Arrange
    cache.put("key1", "value1")

    # Act
    cache.clear()
    result = len(cache)

    # Assert
    assert result == 0, "Cache clear did not empty the cache."


# Test __contains__ method
@pytest.mark.parametrize(
    "key, value, check_key, expected_result",
    [
        ("key1", "value1", "key1", True),  # ID: Contains-True
        ("key2", "value2", "key3", False),  # ID: Contains-False
    ],
)
def test_contains(key, value, check_key, expected_result):
    cache = TimeCache(seconds=10)

    # Arrange
    cache.put(key, value)

    # Act
    result = check_key in cache

    # Assert
    assert result == expected_result, "Cache __contains__ did not return the expected result."


# Test clear_expired method
@pytest.mark.parametrize(
    "key, value, sleep_time, expected_len_after_clear",
    [
        ("key1", "value1", 2, 0),  # ID: ClearExpired-AllExpired
        ("key2", "value2", 0, 1),  # ID: ClearExpired-NoneExpired
    ],
    ids=["AllExpired", "NoneExpired"],
)
def test_clear_expired(key, value, sleep_time, expected_len_after_clear):
    cache = TimeCache(seconds=1)

    # Arrange
    cache.put(key, value)
    if sleep_time:
        time.sleep(sleep_time)

    # Act
    cache.clear_expired()
    result = len(cache)

    # Assert
    assert result == expected_len_after_clear, "Cache clear_expired did not result in the expected cache size."


# Test seconds property
@pytest.mark.parametrize(
    "initial_seconds, new_seconds, expected_seconds",
    [
        (10, 20, 20),  # ID: Seconds-Update
        (1, 0.5, 0.5),  # ID: Seconds-Decrease
    ],
)
def test_seconds_property(initial_seconds, new_seconds, expected_seconds):
    cache = TimeCache(seconds=initial_seconds)

    # Arrange

    # Act
    cache.seconds = new_seconds

    # Assert
    assert cache.seconds == expected_seconds, "Cache seconds property did not update to the expected value."


@pytest.mark.parametrize(
    "seconds, action, key, value, expected_exception",
    [
        (10, "put", None, "A", TypeError),  # ID: ERR-1
        (10, "get", None, None, TypeError),  # ID: ERR-2
        (10, "remove", None, None, TypeError),  # ID: ERR-3
    ],
    ids=["ERR-1", "ERR-2", "ERR-3"],
)
def test_lru_cache_error_cases(seconds, action, key, value, expected_exception):
    # Arrange
    cache = TimeCache(seconds)

    # Act / Assert
    with pytest.raises(expected_exception):
        if action == "put":
            cache.put(key, value)
        elif action == "get":
            cache.get(key)
        elif action == "remove":
            cache.remove(key)
