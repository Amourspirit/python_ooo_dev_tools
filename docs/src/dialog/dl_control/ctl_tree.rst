Class CtlTree
=============

Introduction
------------

Class for working with tree controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(self, src: Any, event: EventArgs, control_src: CtlTree, *args, **kwargs) -> None:
        pass

or

.. code-block:: python

    def on_some_event(self, src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlTree, kwargs['control_src'])

.. autoclass:: ooodev.dialog.dl_control.ctl_tree.CtlTree
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: