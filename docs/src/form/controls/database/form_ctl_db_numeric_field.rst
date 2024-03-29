Class FormCtlDbNumericField
===========================

Introduction
------------

Class for working with Database Numeric Field controls in a form.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: FormCtlDbNumericField, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(FormCtlDbNumericField, kwargs['control_src'])

Class
-----

.. autoclass:: ooodev.form.controls.database.FormCtlDbNumericField
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
