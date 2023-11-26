Class CtlTabPage
================

Introduction
------------

Class for working with tab page controls in a dialog.

EventArgsCallbackT
------------------

All :py:protocol:`~ooodev.utils.type_var.EventArgsCallbackT` callbacks include ``control_src`` as a keyword argument.

A callback can be in the format of:

.. code-block:: python

    def on_some_event(
        src: Any, event: EventArgs, control_src: CtlTabPage, *args, **kwargs
    ) -> None:
        pass

or

.. code-block:: python

    def on_some_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
        # can get control from kwargs
        ctl = cast(CtlTabPage, kwargs['control_src'])

Example
-------

.. cssclass:: screen_shot

    .. image:: https://user-images.githubusercontent.com/4193389/283562180-33f4293a-408d-43d8-92db-d287a4168050.png
        :alt: List Box Examples Screen Shot
        :align: center
        :width: 550px

| For an example see `Tab and List Box Dialog Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/tabs_list_box>`__
| For an example see `Tab and Tree Dialog Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/tree>`__

Class
-----

.. autoclass:: ooodev.dialog.dl_control.CtlTabPage
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: