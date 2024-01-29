.. _ooodev.calc.CalcForm:

Class CalcForm
==============

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

The ``CalcForm`` class is can manage a form for a Calc Sheet.

Adding Controls
---------------

This class has many methods for adding controls to the form. that start with ``insert_control_``
for standard controls and ``insert_db_control_`` for database controls.

Here is an example of adding a button to a form and adding an event handler for the button.

.. code-block:: python

    >>> from typing import Any
    >>> from ooodev.calc import Calc, CalcDoc
    >>> from ooodev.events.args.event_args import EventArgs
    >>> from ooodev.form.controls import FormCtlButton
    >>>
    >>> doc = CalcDoc(Calc.open_doc("form.ods"))
    >>> doc.set_visible()
    >>> sheet = doc.sheets[0]
    >>> if len(sheet.draw_page.forms) == 0:
    ...     sheet.draw_page.forms.add() # add a form with a default name of Form1
    >>> frm = sheet.draw_page.forms[0]
    >>> print(frm.name)
    Form1
    >>> btn = frm.insert_control_button(x=10, y=10, width=40, height=10, label="Button Test")
    >>> btn.add_event_action_performed(on_btn_action_preformed)
    >>>
    >>> def on_btn_action_preformed(src: Any, event: EventArgs, control_src: FormCtlButton, *args, **kwargs) -> None:
    ...    print(
    ...        f"Action Performed: '{control_src.model.Label}', Control Name: {control_src.name}"
    ...    )

Class Declaration
-----------------

.. autoclass:: ooodev.calc.CalcForm
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members: