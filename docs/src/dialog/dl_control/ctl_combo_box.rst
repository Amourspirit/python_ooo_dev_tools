Class CtlComboBox
=================

Introduction
------------

Class for working with combo box controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlComboBox, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlComboBox, kwargs['control_src'])

Class
-----

.. autoclass:: ooodev.dialog.dl_control.CtlComboBox
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: