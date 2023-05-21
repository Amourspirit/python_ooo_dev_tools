class ClassPropertyReadonly(property):
    """
    Class Property. Use with ``@classmethod``.

    .. code-block:: python

        class C(object):
            @ClassProperty
            @classmethod
            def x(cls) -> int:
                return 1

        print(C.x)
        print(C().x)

    NOte:
        This class doesn't actually work for setters, only getters.
    """

    # https://stackoverflow.com/a/1383402/1171746
    def __get__(self, cls, owner):
        return self.fget.__get__(None, owner)()  # type: ignore
