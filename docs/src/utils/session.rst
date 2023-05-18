Class Session
=============

.. autoclass:: ooodev.utils.session.Session
    :members:
    :undoc-members:

    .. py:property:: Session.path_sub

        Static read-only property

        Gets ``com.sun.star.util.PathSubstitution`` instance

        :return: PathSubstitution
        :rtype: str

    .. py:property:: Session.share

        Static read-only property

        Gets Program Share dir,
        such as ``C:\Program Files\LibreOffice\share``

        :return: directory
        :rtype: str

    .. py:property:: Session.shared_scripts

        Static read-only property

        Gets Program Share scripts dir,
        such as ``C:\Program Files\LibreOffice\share\Scripts``

        :return: directory
        :rtype: str

    .. py:property:: Session.shared_py_scripts

        Static read-only property

        Gets Program Share python dir,
        such as ``C:\Program Files\LibreOffice\share\Scripts\python``

        :return: directory
        :rtype: str

    .. py:property:: Session.user_name

        Static read-only property

        Get the username from the environment or password database.

        First try various environment variables, then the password
        database.  This works on Windows as long as USERNAME is set.

        :return: user name
        :rtype: str

    .. py:property:: Session.user_profile

        Static read-only property

        Gets path to user profile such as,
        ``C:\Users\user\AppData\Roaming\LibreOffice\4\user``

        :return: directory
        :rtype: str

    .. py:property:: Session.user_scripts

        Static read-only property

        Gets path to user profile scripts such as,
        ``C:\Users\user\AppData\Roaming\LibreOffice\4\user\Scripts``

        :return: directory
        :rtype: str

    .. py:property:: Session.user_py_scripts

        Static read-only property

        Gets path to user profile python such as,
        ``C:\Users\user\AppData\Roaming\LibreOffice\4\user\Scripts\python``

        :return: directory
        :rtype: str


.. autoclass:: ooodev.utils.session.PathKind
    :members:
    :undoc-members: