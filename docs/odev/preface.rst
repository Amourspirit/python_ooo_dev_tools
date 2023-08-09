*******
Preface
*******

.. image:: https://user-images.githubusercontent.com/4193389/177222844-cf92599a-2fca-4d67-931b-e237f04a3817.png
    :alt: Header Image

Scripting support for any programming language is implemented in LibreOffice and OpenOffice using respective UNO (Universal Network Objects) bridges.
This means that other supported languages, including python, must negotiate a UNO bridge.

This `comment <https://stackoverflow.com/a/64517979/>`__ from Dan Dascalescu offers a challenge to the community:

    after 20 years of software development, the LibreOffice API is the crappiest one I've had the "pleasure" of working with.
    The documentation is `horrible <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html>`__
    ,spread all over the place, `littered with Uyghur <https://bug-attachments.documentfoundation.org/attachment.cgi?id=166685>`__
    , or `completely missing <https://ask.libreoffice.org/en/question/98257/javascript-macro-reference>`__.
    The LibreOffice macro IDE is also extremely unhelpful.

|app_name_bold| (|odev|) offers a python framework for the community to take another step to change this.
The starting point is Java LibreOffice Programming which we thank
Andrew Davison for kindly making available provided his work is acknowledged.

The website the Java code was originally found on is no longer available.
|jlp|_ still has a copy of the website for now. An archive of the Java code is available at `<https://github.com/Amourspirit/libreoffice_lop_java>`__.
The origin book by Andrew Davison can be found in the `LibreOffice Programming <https://flywire.github.io/lo-p/>`__ that was converted by `flywire <https://github.com/flywire>`__.


Changes required are reformatting for this media since the scripts have missed a lot of detail including code,
generalizing the content so it is more widely applicable to supported languages and complementing Python examples with examples of other language.
Please contribute where you can.

.. seealso::

    |lo_dev_guide|_

.. include:: ../resources/odev/authors.rst

.. |lo_dev_guide| replace:: LibreOffice Developer's Guide: Chapter 2 - Professional UNO
.. _lo_dev_guide: https://wiki.documentfoundation.org/Documentation/DevGuide/Professional_UNO

.. |jlp| replace:: WayBack Machine - Java LibreOffice Programming
.. _jlp: https://web.archive.org/web/20221221230318/https://fivedots.coe.psu.ac.th/~ad/jlop/