.. _ooodev.utils.context.lo_context.LoContext:

Class LoContext
===============

The LoContext class is a context manager that provides a way to use multiple documents in |odev|.

Classes in ``ooodev.Draw``, ``ooodev.Calc``, ``ooodev.Writer`` namespaces already use this context manager to enable the use of multiple documents.

In some cases you may want to use muli-documents in your own code. In this case you can use the LoContext class directly.

Only static classes such as :ref:`ooodev.loader.Lo`, :ref:`class_utils_info_info`, :ref:`gui_gui`, and classes in the :ref:`ns_office` namespace would
require the use of the ``LoContext`` class in multi-document mode. Methods and properties of these classes marked as ``|lo_safe|`` or ``|lo_unsafe|``.
This makes them easy to identify. If marked as ``|lo_safe|`` they can be used in multi-document mode without the need to use the ``LoContext`` class.

An example of use is shown below:

.. code-block:: python

    from ooodev.utils.info import Info
    from ooodev.write import WriteDoc
    from ooodev.utils.context import LoContext

    # create first document
    doc1  = WriteDoc.create_doc()

    # create a second document
    lo_inst = Lo.create_lo_instance()
    doc2  = WriteDoc.create_doc(lo_inst=lo_inst)

    # get fonts of first document
    # the default context, which is for the first document is used.
    fonts =  Info.get_fonts()

    # get fonts of second document
    With LoContext(lo_inst=doc2.lo_inst):
        # The context is changed to the second document.
        # As soon as the LoContext block is exited, the context is
        # restored to the first document.
        fonts =  Info.get_fonts()

.. seealso::

    :py:meth:`ooodev.utils.info.Info.get_fonts`


Class Declaration
-----------------

.. autoclass:: ooodev.utils.context.LoContext
    :members: