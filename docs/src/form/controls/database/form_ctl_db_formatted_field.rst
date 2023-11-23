Class FormCtlDbFormattedField
=============================

Introduction
------------

Class for working with Database Formatted Field controls in a form.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlButton, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlButton, kwargs['control_src'])

Class
-----

.. autoclass:: ooodev.form.controls.database.form_ctl_db_formatted_field.FormCtlDbFormattedField
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
