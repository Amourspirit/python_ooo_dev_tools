.. _conn_connect_socket:

Class ConnectSocket
===================

Configuration class for connecting to LibreOffice via socket.
By default no configuration is required.

:py:class:`.Lo.ConnectSocket` is an alias of this class.

Example:

    .. code-block:: python

        from ooodev.loader.lo import Lo

        loader = Lo.load_office(Lo.ConnectSocket())

.. seealso:: 

    - :ref:`ch02`
    - `Starting the LibreOffice Software With Parameters <https://help.libreoffice.org/Common/Starting_the_Software_With_Parameters>`_

.. autoclass:: ooodev.conn.connectors.ConnectSocket
    :members:
    :inherited-members:
    :show-inheritance:
