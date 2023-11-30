Class CtlRadioButton
====================

Introduction
------------

Class for working with radio controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlRadioButton, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlRadioButton, kwargs['control_src'])

Example
-------

.. cssclass:: screen_shot

    .. image:: https://user-images.githubusercontent.com/4193389/283243452-94e5910a-86fb-4d45-ad47-2cb21b266ac4.png
        :alt: Radio Button Example Screen Shot
        :align: center

For an example see `Tab and List Box Dialog Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/grid>`__

Class
-----

.. autoclass:: ooodev.dialog.dl_control.CtlRadioButton
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: