Class CtlGrid
=============

Introduction
------------

Class for working with gird (table) controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlGrid, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlGrid, kwargs['control_src'])

Example
-------

.. cssclass:: screen_shot

    .. image:: https://user-images.githubusercontent.com/4193389/283202594-d22eaa6e-7c27-470e-9c61-f1c94952b903.png
        :alt: Grid Example Screen Shot
        :align: center
        :width: 550px

For an example see `Tab and List Box Dialog Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/grid>`__


Class
-----

.. autoclass:: ooodev.dialog.dl_control.ctl_grid.CtlGrid
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: