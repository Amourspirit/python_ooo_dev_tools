Class CtlPatternField
=====================

Introduction
------------

Class for working with pattern field controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlPatternField, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlPatternField, kwargs['control_src'])

Class
-----

.. autoclass:: ooodev.dialog.dl_control.CtlPatternField
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: