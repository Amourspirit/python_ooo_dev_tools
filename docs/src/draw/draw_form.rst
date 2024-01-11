.. _class_draw_draw_form:

Class DrawForm
===============

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The ``DrawForm`` class is can manage a form for a Write document.

Adding Controls
---------------

This class has many methods for adding controls to the form. that start with ``insert_control_``
for standard controls and ``insert_db_control_`` for database controls.

Here is an example of adding a button to a form and adding an event handler for the button.

.. code-block:: python

    >>> from typing import Any
    >>> from ooodev.draw import Draw, DrawDoc
    >>> from ooodev.events.args.event_args import EventArgs
    >>> from ooodev.form.controls import FormCtlButton
    >>>
    >>> doc =DrawDoc(Draw.create_draw_doc())
    >>> doc.set_visible()
    >>> draw_page = doc.slides[0]
    >>> frm = draw_page.forms.add_form("MainForm")
    >>> print(frm.name)
    MainForm
    >>> btn = frm.insert_control_button(x=10, y=10, width=40, height=10, label="Button Test")
    >>> btn.add_event_action_performed(on_btn_action_preformed)
    >>>
    >>> def on_btn_action_preformed(src: Any, event: EventArgs, control_src: FormCtlButton, *args, **kwargs) -> None:
    ...    print(
    ...        f"Action Performed: '{control_src.model.Label}', Control Name: {control_src.name}"
    ...    )

Class Declaration
-----------------

.. autoclass:: ooodev.draw.DrawForm
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: