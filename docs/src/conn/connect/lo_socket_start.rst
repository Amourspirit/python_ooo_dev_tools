Class LoSocketStart
===================

Class that makes a connection to LibreOffice via socket.

Example:
    .. code-block:: python

        from ooodev.conn.connect import LoSocketStart
        from ooodev.conn.connectors import ConnectSocket
        conn = LoSocketStart(ConnectSocket())
        conn.connect() # may take a few seconds
        smgr = conn.ctx.getServiceManager()
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", conn.ctx)
        # do some work ...
        conn.kill_soffice()

Generally speaking this class will not be called directly. It is used internally by :py:class:`~.utils.lo.Lo`.

Lo Connect Example:

    .. include:: ../../../resources/utils/lo_connect_socket_ex.rst

.. autoclass:: ooodev.conn.connect.LoSocketStart
    :members:
    :inherited-members:
    :show-inheritance:
