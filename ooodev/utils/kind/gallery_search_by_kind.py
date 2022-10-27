from enum import IntEnum


class SearchByKind(IntEnum):
    """
    Gallery Search by Kind.

    Used to determine search criteria.

    See Also:
        :py:meth:`.Gallery.find_gallery_item`
    """

    FILE_NAME = 1
    """Match File Name"""
    TITLE = 2
    """Match Title"""
