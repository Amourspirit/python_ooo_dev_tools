Class ConnectPipe
=================

Configuration class for connecting to LibreOffice via pipe.
By default no configuration is required.

:py:class:`.Lo.ConnectPipe` is an alias of this class.

Example:

    .. code-block:: python

        from ooodev.utils.lo import Lo

        loader = Lo.load_office(Lo.ConnectPipe())

.. seealso::

    - :ref:`ch02`
    - `Starting the LibreOffice Software With Parameters <https://help.libreoffice.org/Common/Starting_the_Software_With_Parameters>`_

.. autoclass:: ooodev.conn.connectors.ConnectPipe
    :members:
    :inherited-members:
    :show-inheritance:
