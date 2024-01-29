.. _ooodev.write.WriteForm:

Class WriteForm
===============

.. image:: https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/75fc2796-6e6b-43e9-b5d1-cf974b8b630f
    :align: center
    :alt: Writer Form
    :width: 500
    :height: 504

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The ``WriteForm`` class is can manage a form for a Write document.

Adding Controls
---------------

This class has many methods for adding controls to the form. that start with ``insert_control_``
for standard controls and ``insert_db_control_`` for database controls.

Here is an example of adding a button to a form and adding an event handler for the button.

.. code-block:: python

    >>> from typing import Any
    >>> from ooodev.write import Write, WriteDoc
    >>> from ooodev.events.args.event_args import EventArgs
    >>> from ooodev.form.controls import FormCtlButton
    >>>
    >>> doc = WriteDoc(Write.create_doc())
    >>> doc.set_visible()
    >>> frm = doc.draw_page.forms.add_form("MainForm")
    >>> print(frm.name)
    MainForm
    >>> btn = frm.insert_control_button(x=10, y=10, width=40, height=10, label="Button Test")
    >>> btn.add_event_action_performed(on_btn_action_preformed)
    >>>
    >>> def on_btn_action_preformed(
    ...     src: Any, event: EventArgs, control_src: FormCtlButton, *args, **kwargs
    ... ) -> None:
    ...     print(
    ...         f"Action Performed: '{control_src.model.Label}', Control Name: {control_src.name}"
    ...     )

Other Examples
--------------

An example can be found on `Live LibreOffice Python UNO Examples <https://github.com/Amourspirit/python-ooouno-ex>`__.
Example: `Build Forms <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_form2>`__.


Class Declaration
-----------------

.. autoclass:: ooodev.write.WriteForm
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: