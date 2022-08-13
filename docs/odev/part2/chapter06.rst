.. _ch06:

**********************
Chapter 6. Text Styles
**********************

.. topic:: Overview

    Five Style Families; Properties; Listing Styles; Creating a Style; Applying Styles;
    Paragraph/Word Styles; Hyperlink Styling; Text Numbering; Headers and Footers

This chapter focuses on how text documents styles can be examined and manipulated.
This revolves around the XStyleFamiliesSupplier_ interface in GenericTextDocument_, which is highlighted in
:numref:`ch06fig_txt_doc_serv_interfaces` (a repeat of :numref:`ch05fig_txt_doc_serv_interfaces` in :ref:`ch05`).

.. cssclass:: diagram invert

    .. _ch06fig_txt_doc_serv_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/181575340-96fb7e21-4e0f-4662-8ed9-92edfb036b0c.png
        :alt: Diagram of The Text Document Services, and some Interfaces

        :The Text Document Services, and some Interfaces.

XStyleFamiliesSupplier_ has a ``getStyleFamilies()`` method for returning text style families.
All these families are stored in an XNameAccess object, as depicted in :numref:`ch06fig_txt_doc_serv_interfaces`.

XNameAccess_ is one of Office's collection types, and employed when the objects in a collection have names.
There's also an XIndexAccess_ for collections in index order.

XNameContainer_ and XIndexContainer_ add the ability to insert and remove objects from a collection.

Work in Progress...

.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _XIndexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XIndexContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexContainer.html
.. _XNameAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameAccess.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XStyleFamiliesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1style_1_1XStyleFamiliesSupplier.html