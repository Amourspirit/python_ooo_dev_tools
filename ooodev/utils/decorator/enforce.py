import sys
import inspect
import typing
from contextlib import suppress
from functools import wraps
from pydoc import locate


def enforce_types(callable):
    """
    Enforces Type base upon the typing value.
    Especially useful but not limited to dataclass
    """
    # https://stackoverflow.com/questions/50563546/validating-detailed-types-in-python-dataclasses
    spec = inspect.getfullargspec(callable)

    def check_types(*args, **kwargs):
        parameters = dict(zip(spec.args, args))
        parameters.update(kwargs)
        for name, value in parameters.items():
            with suppress(KeyError):  # Assume un-annotated parameters can be any type
                type_hint = spec.annotations[name]
                if type(type_hint).__name__ == "str":  # normally should be 'type'
                    # if from __future__ import annotations
                    # then type will be a string.
                    # locate will convert the string to type in most cases
                    # https://stackoverflow.com/questions/11775460/lexical-cast-from-string-to-type
                    type_hint = locate(type_hint)

                if isinstance(type_hint, typing._SpecialForm):
                    # No check for typing.Any, typing.Union, typing.ClassVar (without parameters)
                    continue
                # In Python 3.8 one would replace the try/except with
                # actual_type = typing.get_origin(type_hint) or type_hint
                if sys.version_info >= (3, 8):
                    actual_type = typing.get_origin(type_hint) or type_hint
                else:
                    try:
                        actual_type = type_hint.__origin__
                    except AttributeError:
                        # In case of non-typing types (such as <class 'int'>, for instance)
                        actual_type = type_hint

                if isinstance(actual_type, typing._SpecialForm):
                    # case of typing.Union[…] or typing.ClassVar[…]
                    actual_type = type_hint.__args__  # type: ignore
                is_valid = False
                if isinstance(value, actual_type):  # type: ignore
                    is_valid = True

                if not is_valid:
                    raise TypeError(
                        "Unexpected type for '{}' (expected {} but found {})".format(name, type_hint, type(value))
                    )

    def decorate(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            check_types(*args, **kwargs)
            return func(*args, **kwargs)

        return wrapper

    if inspect.isclass(callable):
        callable.__init__ = decorate(callable.__init__)
        return callable

    return decorate(callable)
