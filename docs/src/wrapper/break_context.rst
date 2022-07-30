.. _wrapper_break_context:

Class BreakContext
===================

``BreakContext`` is a context manager the is designed to wrap other context managers.
If any error other than ``BreakContext.Break`` is raised then ``BreakContext`` will re-raise the error.

If ``BreakContext.Break`` is raised then then the inner context manager exits and
no error is raised. This gives a clean way of exiting a context manager on purpose when
certain conditions are not met.

In this example ``BreakContext`` wraps :py:class:`.Lo.Loader` context manager.
If there is a error opening document then ``BreakContext.Break`` is raised.
:py:class:`.Lo.Loader` terminates the office connection and ``BreakContext``
context manager exits quietly.

.. code-block:: python

    with BreakContext(Lo.Loader(Lo.ConnectSocket(headless=True))) as loader:

        fnm = cast(str, args.file_path)

        try:
            doc = Lo.open_doc(fnm=fnm, loader=loader)
        except Exception:
            print(f"Could not open '{fnm}'")
            raise BreakContext.Break

.. autoclass:: ooodev.wrapper.break_context.BreakContext
    :members: