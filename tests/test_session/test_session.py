import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.session import Session, PathKind


def test_session(loader) -> None:
    import sys

    sub = Session.path_sub
    assert sub is not None
    shared_py = Session.shared_py_scripts
    assert shared_py is not None
    user_py_scripts = Session.user_py_scripts
    assert user_py_scripts is not None
    assert user_py_scripts not in sys.path
    Session.register_path(PathKind.SHARE_USER_PYTHON)
    assert user_py_scripts in sys.path
    sys.path.pop(0)
    assert user_py_scripts not in sys.path

    assert shared_py not in sys.path
    Session.register_path(PathKind.SHARE_PYTHON)
    assert shared_py in sys.path
    sys.path.pop(0)
    assert shared_py not in sys.path
