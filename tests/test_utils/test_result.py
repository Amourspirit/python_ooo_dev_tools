from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.result import Result


def test_result_success() -> None:
    # Test success creation and data access
    result = Result.success(42)
    assert Result.is_success(result)
    assert not Result.is_failure(result)
    assert result.data == 42
    assert result.error is None


def test_result_failure() -> None:
    # Test failure creation and error access
    error = ValueError("test error")
    result = Result.failure(error)
    assert Result.is_failure(result)
    assert not Result.is_success(result)
    assert result.data is None
    assert result.error == error


def test_result_equality() -> None:
    # Test equality comparison
    result1 = Result.success(42)
    result2 = Result.success(42)
    result3 = Result.success(43)
    result4 = Result.failure(ValueError("error"))

    assert result1 == result2
    assert result1 != result3
    assert result1 != result4
    assert result1 != "not a result"


def test_result_unpacking() -> None:
    # Test unpacking functionality
    success_result = Result.success("test")
    failure_result = Result.failure(ValueError("error"))

    # Test unpack() method
    data, error = success_result.unpack()
    assert data == "test"
    assert error is None

    data, error = failure_result.unpack()
    assert data is None
    assert isinstance(error, ValueError)

    # Test iteration unpacking
    data, error = success_result
    assert data == "test"
    assert error is None


def test_result_type_hints() -> None:
    # Test with different types
    str_result = Result.success("hello")
    assert isinstance(str_result.data, str)

    int_result = Result.success(42)
    assert isinstance(int_result.data, int)

    custom_error = TypeError("custom error")
    error_result = Result.failure(custom_error)
    assert isinstance(error_result.error, TypeError)


def test_result_iteration() -> None:
    # Test iteration
    result = Result.success(10)
    values = list(result)
    assert values == [10, None]

    error = ValueError("test error")
    result = Result.failure(error)
    values = list(result)
    assert values == [None, error]


def test_result_complex_types() -> None:
    # Test with more complex types
    data = {"key": "value"}
    result = Result.success(data)
    assert Result.is_success(result)
    assert result.data == data

    class CustomError(Exception):
        pass

    error = CustomError("custom error")
    result = Result.failure(error)
    assert Result.is_failure(result)
    assert isinstance(result.error, CustomError)
