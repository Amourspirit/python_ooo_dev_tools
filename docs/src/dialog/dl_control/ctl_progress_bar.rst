Class CtlProgressBar
====================

Introduction
------------

Class for working with progress bar controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlProgressBar, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlProgressBar, kwargs['control_src'])

Example
-------

.. cssclass:: screen_shot

    .. image:: https://user-images.githubusercontent.com/4193389/284088573-c2bdf23e-7e3a-4ec0-9844-8a70c0be8bdd.png
        :alt: Progress and Scrollbar Screen Shot
        :align: center

For an example see `Progress Bar and Scroll Bar Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/progress_scroll>`__

Class
-----

.. autoclass:: ooodev.dialog.dl_control.CtlProgressBar
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: