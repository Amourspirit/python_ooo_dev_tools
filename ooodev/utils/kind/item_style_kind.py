import uno
from com.sun.star.ui import ItemStyle
from enum import IntFlag


class ItemStyleKind(IntFlag):
    """
    Enum of Const Class ItemStyle

    Specifies styles which influence the appearance and the behavior of an user interface item.

    These styles are only valid if the item describes a toolbar or statusbar item.
    The style values can be combined with the OR operator. Styles which are not valid for an item will be ignored by the implementation.

    To get the int value you can call ``int()`` on the enum value such as ``int(ItemStyleKind.ALIGN_LEFT)``.

    Note:
        This enum is functionally equivalent to the ``ooo.dyn.ui.item_style.ItemStyleEnum`` enum except is include Flags and a ``NONE`` value.
        All enum values are equivalent to the ``com.sun.star.ui.ItemStyle`` constants.

        There are two styles where only one value is valid:

        Alignment:

        - ``ItemStyleKind.ALIGN_LEFT``
        - ``ItemStyleKind.ALIGN_CENTER``
        - ``ItemStyleKind.ALIGN_RIGHT``

        Drawing:

        - ``ItemStyleKind.DRAW_OUT3D``
        - ``ItemStyleKind.DRAW_IN3D``
        - ``ItemStyleKind.DRAW_FLAT``
    """

    NONE = 0
    """None"""

    ALIGN_LEFT = 1
    """
    Specifies how the output of the item is aligned in the bounding box of the user interface element.
    
    This style is only valid for an item which describes a statusbar item. Draw item with a left aligned output.
    """
    ALIGN_CENTER = 2
    """
    Specifies how the output of the item is aligned in the bounding box of the user interface element.
    
    This style is only valid for an item which describes a statusbar item. Draw item with a centered aligned output.
    """
    ALIGN_RIGHT = 3
    """
    Specifies how the output of the item is aligned in the bounding box of the user interface element.
    
    This style is only valid for an item which describes a statusbar item. Draw item with a right aligned output.
    """
    DRAW_OUT3D = 4
    """
    Specifies how the implementation should draw the item.
    
    This style is only valid for an item which describes a statusbar item. Draw item with an embossed 3D effect.
    """
    DRAW_IN3D = 8
    """
    Specifies how the implementation should draw the item.
    
    This style is only valid for an item which describes a statusbar item. Draw item with an impressed 3D effect.
    """
    DRAW_FLAT = 12
    """
    Specifies how the implementation should draw the item.
    
    This style is only valid for an item which describes a statusbar item. Draw item without an 3D effect.
    """
    OWNER_DRAW = 16
    """
    Specifies whether or not an item is displayed using an external function.
    
    This style is only valid if the item describes a statusbar item.
    """
    AUTO_SIZE = 32
    """
    Specifies whether or not the size of the item is set automatically by the parent user interface element.
    
    This style is only valid if the item describes a toolbar or statusbar item.
    """
    RADIO_CHECK = 64
    """
    Determines whether the item unchecks neighbor entries which have also this style set.
    
    This style is only valid if the item describes a toolbar item.
    """
    ICON = 128
    """
    Specifies if an icon is placed on left side of the text, like an entry in a taskbar.
    
    This style is only valid if the item describes a toolbar item and visible if style of the toolbar is set to symboltext.
    
    This style can also be used for custom toolbars and menus, in a custom toolbar an item's Style setting can used to override the toolbar container setting, the style can be bitwise OR-ed with com.sun.star.ui.ItemStyle.TEXT to define text, text+icon or icon only is to be displayed. Similarly for menu items, an items Style can override the application setting to display either text or icon (note: for menu an icon only setting interpreted as icon+text)
    """
    DROP_DOWN = 256
    """
    Specifies that the item supports a dropdown menu or toolbar for additional functions.
    
    This style is only valid if the item describes a toolbar item.
    """
    REPEAT = 512
    """
    Indicates that the item continues to execute the command while you click and hold the mouse button.
    
    This style is only valid if the item describes a toolbar item.
    """
    DROPDOWN_ONLY = 1024
    """
    Indicates that the item only supports a dropdown menu or toolbar for additional functions.
    
    There is no function on the button itself.
    
    This style is only valid if the item describes a toolbar item.
    """
    TEXT = 2048
    """
    indicates if icon, text or text+icon is displayed for the item.
    
    This style can be used for custom toolbars and menus, in a custom toolbar an item's Style setting can used to override the toolbar container setting, the style can be bitwise OR-ed with com.sun.star.ui.ItemStyle.ICON to define text, text+icon or icon only is to be displayed. Similarly for menu items, an items Style can override the application setting to display either text or icon (note: for menu an icon only setting interpreted as icon+text)
    """
    MANDATORY = 4096
    """
    Marks always visible element which can not be removed when statusbar width is not sufficient.
    
    **since**
    
        LibreOffice 6.1
    """

    def to_json(self) -> int:
        return int(self)
