# coding: utf-8
"""
Generic Context managers
"""


class BreakContext(object):
    """
    Generic Context manager that allow for BreakContext. Break statement to exit with block.

    Example:

        .. code-block:: python

            with BreakContext(open(path)) as f:
                print 'before condition'
                if condition:
                    raise BreakContext.Break
                print 'after condition'
    """

    class Break(Exception):
        """Break out of the with statement"""

        ...
        # https://stackoverflow.com/questions/11195140/break-or-exit-out-of-with-statement

    def __init__(self, value):
        self.value = value

    def __enter__(self):
        return self.value.__enter__()

    def __exit__(self, etype, value, traceback):
        error = self.value.__exit__(etype, value, traceback)
        if etype == self.Break:
            return True
        return error
