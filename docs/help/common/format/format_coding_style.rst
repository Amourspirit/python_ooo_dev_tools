.. _help_format_coding_style:

Format Coding Style
===================

The :py:mod:`~ooodev.format` module constains hundreds of classes.
The following coding style is used to make the code more usable.

Properties
----------

All Properties that belong to the style are prefixed with ``prop_``.

This means any other property names return an new instance of the style class with the property set.

For example:

``base_font`` is a new instance of the :py:class:`~ooodev.format.writer.direct.char.font.Font` class set to color ``StandardColor.GREEN_DARK3``, size ``12`` and font ``Liberation Serif``.

``base_font.italic`` when applied to :py:meth:`~.write.Write.append_para` method  is called using ``base_font.italic.bold``
that will return a new instance of the :py:class:`~ooodev.format.writer.direct.char.font.Font` class with the ``italic`` and ``bold`` properties set.
This will result in a font being applied with the following properties:

    * color ``StandardColor.GREEN_DARK3``
    * size ``12``
    * font ``Liberation Serif``
    * italic ``True``
    * bold ``True``

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 6

        from ooodev.format.writer.direct.char.font import Font
        from ooodev.utils.color import StandardColor
        # ... other code

        base_font = Font(color=StandardColor.GREEN_DARK3, size=12, name="Liberation Serif")
        Write.append_para(cursor=cursor, text="Hello World!", styles=[base_font.italic.bold])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Methods
-------

apply
+++++

The ``apply()`` method is used to apply the style to the document object.
Depending on the style, it may be applied to different objects.

For instance :py:class:`ooodev.format.writer.direct.char.font.Font` is applied to Writer Text Characters usually via a cursor or
one of ``Write``'s methods such as :py:meth:`~.Write.append`.

fmt\_ methods
+++++++++++++

Style methods that start with ``fmt_`` prefix usually take a single argument and return a new instance of the style class with the property set.

For example:

``spc_font`` is a new instance :py:class:`~ooodev.format.writer.direct.char.font.Font` with the ``spacing`` value set to ``2.1``
and the property value can be retreived using ``spc_font.prop_spacing``.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 6

        from ooodev.format.writer.direct.char.font import Font
        from ooodev.utils.color import StandardColor
        # ... other code

        base_font = Font(color=StandardColor.GREEN_DARK3, size=12, name="Liberation Serif")
        spc_font = base_font.fmt_spacing(2.1)
        Write.append_para(cursor=cursor, text="Hello World!", styles=[spc_font])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_writer_format_direct_char_font`
