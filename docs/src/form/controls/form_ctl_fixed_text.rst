Class FormCtlFixedText
======================

Introduction
------------

Class for working with Fixed Text controls in a form.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: FormCtlFixedText, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(FormCtlFixedText, kwargs['control_src'])

Class
-----

.. autoclass:: ooodev.form.controls.FormCtlFixedText
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
