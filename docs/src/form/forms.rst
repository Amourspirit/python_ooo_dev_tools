.. _ooodev.form.Forms:

Class Forms
===========

.. image:: https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/75fc2796-6e6b-43e9-b5d1-cf974b8b630f
    :align: center
    :alt: Forms
    :width: 500
    :height: 504

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Introduction
------------

This class is used to create create and manage forms and form controls.

Multi-Document Support
----------------------

Multi-Document support is available in |odev|.
The :ref:`ooodev.draw.DrawForm`, :ref:`ooodev.calc.CalcForm`, and :ref:`ooodev.writer.WriterForm` classes all inherit from
the :py:class:`~ooodev.form.partial.form_partial.FormsPartial` class which has multi document support built in.

When creating your own form class it is recommended it inherit from the :py:class:`~ooodev.form.partial.form_partial.FormsPartial` class.

Methods in the this ``Forms`` class are identified if they have multi-document support or not.

For methods that do no have multi-document support you will need to use the :ref:`ooodev.utils.context.lo_context.LoContext` context manager before calling them.

.. code-block:: python

    from ooodev.form import Forms
    from ooodev.write import WriteDoc
    from ooodev.utils.context.lo_context import LoContext

    # create first document
    doc1  = WriteDoc.create_doc()

    # create a second document
    lo_inst = Lo.create_lo_instance()
    doc2  = WriteDoc.create_doc(lo_inst=lo_inst)

    # ...

    with LoContext(lo_inst=lo_inst):
        # in this block all methods will automatically use
        # the lo_inst of the second document.
        # As soon as the LoContext block is exited, the context is
        # restored to the first document.
        # ...
        Forms.add_control(
            doc=doc2.component,
            name=name,
            label="Options",
            comp_kind=FormComponentKind.GROUP_BOX,
            x=col2_x,
            y=y,
            width=box_width,
            height=25,
            styles=[font_colored],
        )
        # ...
        props = Forms.add_list(
            doc=doc2.component,
            name="Fruits",
            entries=fruits,
            x=x,
            y=y,
            width=width,
            height=height,
        )

        # ...


Examples
--------

A few examples can be found on `Live LibreOffice Python UNO Examples <https://github.com/Amourspirit/python-ooouno-ex>`__.

- Writer `Build Form <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_form>`__
- Writer `Build Form2 <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_form2>`__
- Draw `Build Form <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_build_form>`__

Class Declaration
-----------------
.. autoclass:: ooodev.form.Forms
    :members:
    :undoc-members: