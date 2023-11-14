Class CtlListBox
================

Introduction
------------

Class for working with list box controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(self, src: Any, event: EventArgs, control_src: CtlListBox, *args, **kwargs) -> None:
        pass

or

.. code-block:: python

    def on_some_event(self, src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlListBox, kwargs['control_src'])

.. autoclass:: ooodev.dialog.dl_control.ctl_list_box.CtlListBox
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: