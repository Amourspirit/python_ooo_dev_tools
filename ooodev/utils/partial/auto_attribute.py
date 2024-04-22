class AutoAttribute:
    """
    Automatically add attributes to the object when they are set.
    Useful for debugging and testing.

    Example:

        .. code-block:: python

                obj = AutoAddAttributes()
                obj.new_attribute = "New value"

    .. versionadded:: 0.41.0
    """

    def __setattr__(self, name, value):
        # You can add any custom logic here
        # print(f"Setting attribute {name} to {value}")
        self.__dict__[name] = value
