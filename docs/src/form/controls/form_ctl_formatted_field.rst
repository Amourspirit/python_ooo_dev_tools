Class FormCtlFormattedField
===========================

Introduction
------------

Class for working with Formatted Field controls in a form.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: FormCtlFormattedField, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(FormCtlFormattedField, kwargs['control_src'])

Class
-----

.. autoclass:: ooodev.form.controls.FormCtlFormattedField
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
