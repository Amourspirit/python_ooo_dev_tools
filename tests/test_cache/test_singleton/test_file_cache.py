import pytest
import json
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.cache.singleton import FileCache


@pytest.fixture
def cache(loader):
    return FileCache(tmp_dir="test_file_cache", key="test_key")


def test_put_and_get(cache, tmp_path):
    data = {"name": "John", "age": 30, "city": "New York"}
    # Create a temporary file path
    temp_file = tmp_path / "temp_data.json"

    # Write JSON data to the temporary file
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(data, f)

    # Put content into cache
    cache.put(temp_file, data)
    assert temp_file in cache

    # Get content from cache
    retrieved_content = cache.get(temp_file)

    assert retrieved_content == data

    retrieved_content = cache[temp_file]
    assert retrieved_content == data

    # Get cached content from cache
    retrieved_content = cache.get(temp_file)

    assert retrieved_content == data

    # Get cached content from cache
    retrieved_content = cache[temp_file]

    assert retrieved_content == data


def test_remove(cache, tmp_path):
    data = {"name": "John", "age": 30, "city": "New York"}
    # Create a temporary file path
    temp_file = tmp_path / "temp_data.json"

    # Write JSON data to the temporary file
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(data, f)

    # Put content into cache
    cache.put(temp_file, data)
    assert temp_file in cache

    # Remove content from cache
    cache.remove(temp_file)

    # Try to get content from cache
    retrieved_content = cache.get(temp_file)

    assert retrieved_content is None

    # Put content into cache
    cache[temp_file] = data

    # Try to get content from cache
    retrieved_content = cache[temp_file]
    assert temp_file in cache
    assert retrieved_content == data

    del cache[temp_file]
    retrieved_content = cache[temp_file]
    assert retrieved_content is None


def test_singleton(cache):
    cache2 = FileCache(tmp_dir="test_file_cache", key="test_key")
    assert cache is cache2
    cache3 = FileCache(tmp_dir="test_file_cache")
    assert cache is not cache3


def test_path_init(tmp_path, loader):
    cache = FileCache(tmp_dir=Path("test_txt_cache/nice"), key="test_key")
    data = {"name": "John", "age": 30, "city": "New York"}
    # Create a temporary file path
    temp_file = tmp_path / "temp_data.json"

    # Write JSON data to the temporary file
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(data, f)

    # Put content into cache
    cache.put(temp_file, data)
    assert temp_file in cache


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
