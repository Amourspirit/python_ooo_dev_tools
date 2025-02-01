import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.cache.singleton.lru_cache import LRUCache


@pytest.fixture
def cache():
    return LRUCache(capacity=10, key="test_key")


def test_singleton(cache):
    cache2 = LRUCache(capacity=10, key="test_key")
    assert cache is cache2
    cache3 = LRUCache()
    assert cache is not cache3

    cache4 = LRUCache()
    assert cache3 is cache4

    cache5 = LRUCache(capacity=20)
    cache6 = LRUCache(capacity=20)
    assert cache5 is cache6

    cache7 = LRUCache(capacity=30)
    assert cache6 is not cache7

    cache8 = LRUCache(key="test_key", key2="test_key2", key3=11)
    cache9 = LRUCache(key="test_key", key2="test_key2", key3=11)
    assert cache8 is cache9


# Happy Path Tests
@pytest.mark.parametrize(
    "capacity, actions, expected",
    [
        (
            2,
            [("put", 1, "A"), ("put", 2, "B"), ("get", 1), ("put", 3, "C"), ("get", 2), ("get", 3)],
            ["A", None, "C"],
        ),  # ID: HP-1
        (
            3,
            [
                ("put", "a", 100),
                ("put", "b", 200),
                ("get", "a"),
                ("put", "c", 300),
                ("put", "d", 400),
                ("get", "b"),
                ("get", "c"),
                ("get", "d"),
            ],
            [100, None, 300, 400],
        ),  # ID: HP-2
        (
            1,
            [("put", "x", "X"), ("get", "x"), ("put", "y", "Y"), ("get", "x"), ("get", "y")],
            ["X", None, "Y"],
        ),  # ID: HP-3
    ],
    ids=["HP-1", "HP-2", "HP-3"],
)
def test_lru_cache_happy_path(capacity, actions, expected):
    # Arrange
    cache = LRUCache(capacity=capacity)
    results = []

    # Act
    for action in actions:
        if action[0] == "put":
            cache.put(action[1], action[2])
        elif action[0] == "get":
            results.append(cache.get(action[1]))

    # Assert
    assert results == expected


# Edge Cases
@pytest.mark.parametrize(
    "capacity, actions, expected_len, expected_repr",
    [
        (0, [("put", 1, "A"), ("get", 1)], 0, "LRUCache(0)"),  # ID: EC-1
        (-1, [("put", 2, "B")], 0, "LRUCache(0)"),  # ID: EC-2
        (2, [("put", 1, "A"), ("put", 2, "B"), ("put", 1, "C"), ("put", 3, "D")], 2, "LRUCache(2)"),  # ID: EC-3
    ],
    ids=["EC-1", "EC-2", "EC-3"],
)
def test_lru_cache_edge_cases(capacity, actions, expected_len, expected_repr):
    # Arrange
    cache = LRUCache(capacity=capacity)

    # Act
    for action in actions:
        if action[0] == "put":
            cache.put(action[1], action[2])
        elif action[0] == "get":
            cache.get(action[1])

    # Assert
    assert len(cache) == expected_len
    assert repr(cache) == expected_repr


# Error Cases
@pytest.mark.parametrize(
    "capacity, action, key, value, expected_exception",
    [
        (1, "put", None, "A", TypeError),  # ID: ERR-1
        (1, "get", None, None, TypeError),  # ID: ERR-2
        (1, "remove", None, None, TypeError),  # ID: ERR-3
    ],
    ids=["ERR-1", "ERR-2", "ERR-3"],
)
def test_lru_cache_error_cases(capacity, action, key, value, expected_exception):
    # Arrange
    cache = LRUCache(capacity=capacity)

    # Act / Assert
    with pytest.raises(expected_exception):
        if action == "put":
            cache.put(key, value)
        elif action == "get":
            cache.get(key)
        elif action == "remove":
            cache.remove(key)
