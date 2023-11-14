Class CtlFixedLine
==================

Introduction
------------

Class for working with fixed line controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(self, src: Any, event: EventArgs, control_src: CtlFixedLine, *args, **kwargs) -> None:
        pass

or

.. code-block:: python

    def on_some_event(self, src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlFixedLine, kwargs['control_src'])

.. autoclass:: ooodev.dialog.dl_control.ctl_fixed_line.CtlFixedLine
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: