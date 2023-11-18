Class CtlTextEdit
=================

Introduction
------------

Class for working with text edit controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlTextEdit, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlTextEdit, kwargs['control_src'])

Example
-------

.. cssclass:: screen_shot

    .. image:: https://user-images.githubusercontent.com/4193389/284018833-91fdd4ac-a2c2-4105-a32b-922480240a12.png
        :alt: Tab and Dialog Example
        :align: center

For an example see `Tab and Tree Dialog Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/tree>`__

Class
-----

.. autoclass:: ooodev.dialog.dl_control.ctl_text_edit.CtlTextEdit
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: