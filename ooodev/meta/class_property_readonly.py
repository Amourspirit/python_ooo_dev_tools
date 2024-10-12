class ClassPropertyReadonly:
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

    def __init__(self, fget):
        self.fget = fget

    def __get__(self, instance, owner):
        return self.fget.__get__(None, owner)()
        # return self.fget(owner)
