Class Info
==========

.. autoclass:: ooodev.utils.info.Info
    :members:
    :undoc-members:

    .. py:property:: Info.language

        Static read-only property

        Gets the Current Language of the LibreOffice Instance

        :return: First two chars of language in lower case such as ``en-US``
        :rtype: str
    
    .. py:property:: Info.version

        Static read-only property

        Gets the running LibreOffice version

        :return: version as string such as ``"7.3.4.2"``
        :rtype: str
    
    .. py:property:: Info.version_info

        Static read-only property

        Gets the running LibreOffice version

        :return: version as tuple of ints such as ``(7, 3, 4, 2)``
        :rtype: tuple