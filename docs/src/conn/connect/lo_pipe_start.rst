Class LoPipeStart
=================

Class that makes a connection to LibreOffice via pipe.

Example:
    .. code-block:: python

        from ooodev.conn.connect import LoPipeStart
        from ooodev.conn.connectors import ConnectPipe
        conn = LoPipeStart(ConnectPipe())
        conn.connect() # may take a few seconds
        smgr = conn.ctx.getServiceManager()
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", conn.ctx)
        # do some work ...
        conn.kill_soffice()

Generally speaking this class will not be called directly. It is used internally by :py:class:`~.utils.lo.Lo`.

Lo Connect Example:

    .. include:: ../../../resources/utils/lo_connect_pipe_ex.rst


.. autoclass:: ooodev.conn.connect.LoPipeStart
    :members:
    :inherited-members:
    :show-inheritance:
