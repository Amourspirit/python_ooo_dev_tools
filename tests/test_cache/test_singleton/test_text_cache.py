import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.cache.singleton import TextCache


@pytest.fixture
def cache(loader):
    return TextCache(tmp_dir="test_txt_cache", lifetime=60, key="test_key")


def test_put_and_get(cache):
    data = "Hello World!"
    # Create a temporary file path

    file_name = "temp_data.txt"

    # Put content into cache
    cache.put(file_name, data)
    assert file_name in cache

    # Get content from cache
    retrieved_content = cache.get(file_name)

    assert retrieved_content == data

    retrieved_content = cache[file_name]
    assert retrieved_content == data

    # Get cached content from cache
    retrieved_content = cache.get(file_name)

    assert retrieved_content == data

    # Get cached content from cache
    retrieved_content = cache[file_name]

    assert retrieved_content == data


def test_remove(cache):
    data = "Hello World!"
    # Create a temporary file path
    file_name = "temp_data.txt"

    # Put content into cache
    cache.put(file_name, data)
    assert file_name in cache

    # Remove content from cache
    cache.remove(file_name)

    # Try to get content from cache
    retrieved_content = cache.get(file_name)

    assert retrieved_content is None

    # Put content into cache
    cache[file_name] = data

    # Try to get content from cache
    retrieved_content = cache[file_name]
    assert file_name in cache
    assert retrieved_content == data

    del cache[file_name]
    retrieved_content = cache[file_name]
    assert retrieved_content is None


def test_singleton(cache):
    cache2 = TextCache(tmp_dir="test_txt_cache", lifetime=60, key="test_key")
    assert cache is cache2
    cache3 = TextCache(tmp_dir="test_txt_cache")
    assert cache is not cache3


def test_path_init(loader):
    cache = TextCache(tmp_dir=Path("test_txt_cache/nice"), lifetime=300, key="test_key")
    data = "Hello World!"
    # Create a temporary file path
    file_name = "temp_data.txt"

    # Put content into cache
    cache.put(file_name, data)
    assert file_name in cache


# Error Cases
@pytest.mark.parametrize(
    "action, file_path, value, expected_exception",
    [
        ("put", None, "A", ValueError),  # ID: ERR-1
        ("get", None, "A", ValueError),  # ID: ERR-2
        ("remove", None, "A", ValueError),  # ID: ERR-3
        ("put", "", "A", ValueError),  # ID: ERR-4
        ("get", "", "A", ValueError),  # ID: ERR-5
        ("remove", "", "A", ValueError),  # ID: ERR-6
    ],
    ids=["ERR-1", "ERR-2", "ERR-3", "ERR-4", "ERR-5", "ERR-6"],
)
def test_lru_cache_error_cases(action, file_path, value, expected_exception, cache):
    # Arrange

    # Act / Assert
    with pytest.raises(expected_exception):
        if action == "put":
            cache.put(file_path, value)
        elif action == "get":
            cache.get(file_path)
        elif action == "remove":
            cache.remove(file_path)

    with pytest.raises(expected_exception):
        if action == "put":
            cache[file_path] = value
        elif action == "get":
            cache[file_path]
        elif action == "remove":
            del cache[file_path]
