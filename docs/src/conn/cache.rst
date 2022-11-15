.. _conn_cache:

Class Cache
===========

``Cache`` is often used in conjunction with :py:class:`~.conn.connectors.ConnectPipe` or :py:class:`~.conn.connectors.ConnectSocket`

``Cache`` provides caching option for profiles when connecting to LibreOffice(not used in macros).
In some cases it is beneficial (such as when testing and debugging) to copy as profile that will be used for a single LibreOffice session.

``Cache`` provides a way to supply a preexisting profile and copy that profile into a temp working directory.
When cache is used any changes made in LibreOffice will only last for the session.
Once the session is terminated the working dir is invalidated.

Subsequent sessions will get a fresh profile copied from cache.

``Cache`` default settings searches in know locations for current users profile and creates, copies the profile into a temp dir, then sets the temp dir as working dir.

Example connecting using Cache:

    .. include:: ../../resources/utils/lo_connect_socket_cache_ex.rst

.. seealso:: 

    - :ref:`ch02`

.. autoclass:: ooodev.conn.cache.Cache
    :members:
