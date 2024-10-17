from __future__ import annotations
import datetime
from typing import Any, TYPE_CHECKING, cast, Tuple
from pathlib import Path

from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.awt.font_slant import FontSlant
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum

from ooodev.dialog import ImageScaleModeEnum, BorderKind, DateFormatKind, TimeFormatKind, StateKind
from ooodev.loader import lo as mLo
from ooodev.calc import CalcDoc
from ooodev.utils.kind.orientation_kind import OrientationKind
from ooodev.write import WriteDoc
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.color import StandardColor
from ooodev.dialog.dl_control.ctl_date_field import CtlDateField
from ooodev.dialog import TriStateKind
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.info import Info
from ooodev.units import UnitAppFontHeight

if TYPE_CHECKING:
    from com.sun.star.awt import ItemEvent
    from com.sun.star.awt import AdjustmentEvent
    from com.sun.star.awt import WindowEvent
    from com.sun.star.beans import PropertyChangeEvent
    from ooodev.dialog.dl_control.ctl_button import CtlButton
    from ooodev.dialog.dl_control.ctl_check_box import CtlCheckBox
    from ooodev.dialog.dl_control.ctl_scroll_bar import CtlScrollBar
    from ooodev.dialog.dl_control.ctl_dialog import CtlDialog
    from ooodev.dialog.dl_control.ctl_base import DialogControlBase


class Runner:
    # pylint: disable=unused-argument
    def __init__(
        self,
        doc: WriteDoc | CalcDoc,
        title: str,
        msg: str,
        input_value: str = "",
        ok_lbl: str = "OK",
        cancel_lbl: str = "Cancel",
        is_password: bool = False,
    ) -> None:
        self._doc = doc
        self._init_handlers()

        # self._dialog = cast(
        #     "UnoControlDialog",
        #     mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        # )
        # dialog_model = mLo.Lo.create_instance_mcf(
        #     XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
        # )

        # self._dialog.setModel(dialog_model)
        # Liberation Serif Regular
        # print(Info.get_font_names())
        fd = Info.get_font_descriptor("Liberation Serif", "Regular")
        if fd is not None:
            fd.Height = 10
        print(fd)

        border_kind = BorderKind.BORDER_SIMPLE
        self._title = title
        self._width = 800
        self._height = 700
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 4
        self._box_height = 30
        if border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._tab_index = 1
        self._dialog = self._doc.create_dialog(
            x=-1,
            y=-1,
            width=self._width,
            height=self._height,
            title=self._title,
        )
        # self._dialog.create_peer()
        # self._dialog.set_visible(False)

        self._ctl_lbl = self._dialog.insert_label(
            label=msg,
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=20,
        )
        # print(self._ctl_lbl.font_descriptor.component)
        if fd is not None:
            self._ctl_lbl.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_lbl)
        self._ctl_lbl.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_lbl.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_lbl.view.getPosSize()
        if is_password:
            self._txt_input = self._dialog.insert_password_field(
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
            )
        else:
            self._txt_input = self._dialog.insert_text_field(
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
            )
        self._set_tab_index(self._txt_input)
        if fd is not None:
            self._txt_input.set_font_descriptor(fd)
        self._txt_input.add_event_text_changed(self._fn_on_text_changed)
        self._txt_input.text_color = StandardColor.GREEN_DARK2

        self._ctl_btn_cancel = self._dialog.insert_button(
            label=cancel_lbl,
            x=self._width - self._btn_width - self._margin,
            y=self._height - self._btn_height - self._margin,
            width=self._btn_width,
            height=self._btn_height,
            # btn_type=PushButtonType.CANCEL,
        )
        if fd is not None:
            self._ctl_btn_cancel.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_btn_cancel)
        self._ctl_btn_cancel.view.setActionCommand("Cancel")
        self._ctl_btn_cancel.add_event_action_performed(self._fn_on_action_cancel)
        self._ctl_btn_cancel.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_btn_cancel.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_btn_cancel.view.getPosSize()
        self._ctl_button_ok = self._dialog.insert_button(
            label=ok_lbl,
            x=sz.X - sz.Width - self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )
        if fd is not None:
            self._ctl_button_ok.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_button_ok)
        self._ctl_button_ok.add_event_action_performed(self._fn_on_action_ok)
        self._ctl_button_ok.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_button_ok.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._txt_input.view.getPosSize()
        # ctl_button_ok.width += 30
        # ctl_button_ok.x -= 30
        self._ctl_chk1 = self._dialog.insert_check_box(
            label="Check Box 1",
            x=sz.X,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=UnitAppFontHeight(11),
            tri_state=False,
            state=TriStateKind.CHECKED,
            border=border_kind,
        )
        if fd is not None:
            self._ctl_chk1.set_font_descriptor(fd)
        # self._ctl_chk1.model_checkbox.x = self._ctl_chk1.x.get_value_app_font(0)
        # self._ctl_chk1.model_checkbox.y = self._ctl_chk1.y.get_value_app_font(1)
        # self._ctl_chk1.model_checkbox.width = self._ctl_chk1.width.get_value_app_font(2)
        # self._ctl_chk1.model_checkbox.height = self._ctl_chk1.height.get_value_app_font(3)
        self._ctl_chk1.text_color = StandardColor.RED
        self._set_tab_index(self._ctl_chk1)

        sz = self._ctl_chk1.view.getPosSize()
        self._ctl_chk2 = self._dialog.insert_check_box(
            label="Check Box 2",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=False,
            state=TriStateKind.NOT_CHECKED,
            border=border_kind,
        )
        if fd is not None:
            self._ctl_chk2.set_font_descriptor(fd)
        self._ctl_chk2.text_color = StandardColor.GREEN
        self._set_tab_index(self._ctl_chk2)

        sz = self._ctl_chk2.view.getPosSize()
        self._ctl_chk3 = self._dialog.insert_check_box(
            label="Check Box 3",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=True,
            state=TriStateKind.DONT_KNOW,
            border=border_kind,
        )
        if fd is not None:
            self._ctl_chk3.set_font_descriptor(fd)
        self._ctl_chk3.text_color = StandardColor.BLUE
        self._set_tab_index(self._ctl_chk3)
        self._ctl_chk1.add_event_item_state_changed(self._fn_on_check_box_state)
        self._ctl_chk2.add_event_item_state_changed(self._fn_on_check_box_state)
        self._ctl_chk3.add_event_item_state_changed(self._fn_on_check_box_state)
        self._ctl_chk1.font_descriptor.weight = 75
        self._ctl_chk2.font_descriptor.weight = 75
        self._ctl_chk3.font_descriptor.weight = 75

        sz = self._ctl_chk1.view.getPosSize()
        self._ctl_date = self._dialog.insert_date_field(
            x=sz.Width + self._padding,
            y=sz.Y,
            width=190,
            height=self._box_height,
            date_value=datetime.datetime.now(),
            border=border_kind,
        )
        if fd is not None:
            self._ctl_date.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_date)
        self._ctl_date.date_format = DateFormatKind.DIN_5008_YY_MM_DD
        dt = datetime.datetime.now()
        self._ctl_date.date = datetime.datetime(dt.year - 1, dt.month, dt.day)

        self._ctl_date.add_event_down(self._fn_on_down)
        self._ctl_date.add_event_up(self._fn_on_up)
        self._ctl_date.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_date.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_date.view.getPosSize()
        self._ctl_currency = self._dialog.insert_currency_field(
            x=sz.Width + sz.X + self._padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            value=123.45,
            spin_button=True,
            border=border_kind,
        )
        if fd is not None:
            self._ctl_currency.set_font_descriptor(fd)
        self._ctl_currency.font_descriptor.height = 10
        self._ctl_currency.text_color = StandardColor.BLUE
        self._set_tab_index(self._ctl_currency)
        sz = self._ctl_currency.view.getPosSize()
        self._ctl_currency.add_event_down(self._fn_on_down)
        self._ctl_currency.add_event_up(self._fn_on_up)
        self._ctl_pattern = self._dialog.insert_pattern_field(
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=sz.Width,
            height=sz.Height,
            edit_mask="NNLNNLLLLL",
            literal_mask="__.__.2025",
            border=border_kind,
        )
        if fd is not None:
            self._ctl_pattern.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_pattern)
        self._ctl_pattern.add_event_down(self._fn_on_down)
        self._ctl_pattern.add_event_up(self._fn_on_up)
        self._ctl_pattern.add_event_text_changed(self._fn_on_text_changed)

        sz_date = self._ctl_date.view.getPosSize()
        self._ctl_num_field = self._dialog.insert_numeric_field(
            x=sz_date.X,
            y=sz_date.Y + sz_date.Height + self._padding,
            width=sz.Width,
            height=self._box_height,
            value=123,
            spin_button=True,
            border=border_kind,
        )
        if fd is not None:
            self._ctl_num_field.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_num_field)
        self._ctl_num_field.add_event_down(self._fn_on_down)
        self._ctl_num_field.add_event_up(self._fn_on_up)
        self._ctl_num_field.add_event_text_changed(self._fn_on_text_changed)

        sz_numeric = self._ctl_num_field.view.getPosSize()
        # sz_fmt = ctl_formatted.getPosSize()
        self._ctl_combo1 = self._dialog.insert_combo_box(
            x=self._margin,
            y=sz_numeric.Height + sz_numeric.Y + self._padding,
            width=200,
            height=self._box_height,
            entries=["Item 1", "Item 2", "Item 3"],
            border=border_kind,
            drop_down=True,
        )
        # self._ctl_combo1 = self._dialog.insert_list_box(
        #     x=self._margin,
        #     y=sz_numeric.Height + sz_numeric.Y + self._padding,
        #     width=200,
        #     height=self._box_height,
        #     entries=["Item 1", "Item 2", "Item 3"],
        #     border=border_kind,
        #     drop_down=True,
        # )
        if fd is not None:
            self._ctl_combo1.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_combo1)
        self._ctl_combo1.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_combo1.add_event_item_state_changed(self._fn_on_item_changed)

        sz = self._ctl_combo1.view.getPosSize()

        self._ctl_progress = self._dialog.insert_progress_bar(
            x=sz_date.X,
            y=sz.Y,
            width=400,
            height=self._box_height,
            min_value=1,
            value=67,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_progress)
        self._ctl_progress.fill_color = StandardColor.GREEN
        self._ctl_progress.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_progress.add_event_mouse_exited(self._fn_on_mouse_exit)

        self._ctl_scroll_progress = self._dialog.insert_scroll_bar(
            x=self._ctl_progress.x,
            y=self._ctl_progress.y + self._ctl_progress.height + self._padding,
            width=self._ctl_progress.width,
            height=self._box_height,
            min_value=self._ctl_progress.model.ProgressValueMin,
            max_value=self._ctl_progress.model.ProgressValueMax,
        )
        self._set_tab_index(self._ctl_scroll_progress)
        self._ctl_scroll_progress.value = self._ctl_progress.value
        self._ctl_scroll_progress.add_event_adjustment_value_changed(self._fn_on_scroll_adjustment)

        self._ctl_file = self._dialog.insert_file_control(
            x=sz.X,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=self._box_height,
            border=border_kind,
        )
        if fd is not None:
            self._ctl_file.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_file)
        self._ctl_file.text = "file:///workspace/ooouno-dev-tools/tests/fixtures/image/img_brick.png"
        self._ctl_file.text_color = StandardColor.BLUE_LIGHT3
        self._ctl_file.add_event_text_changed(self._fn_on_text_changed)
        sz = self._ctl_file.view.getPosSize()
        self._ctl_ln = self._dialog.insert_fixed_line(
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=self._width - (self._margin * 2),
            height=1,
        )
        self._set_tab_index(self._ctl_ln)

        sz = self._ctl_ln.view.getPosSize()
        self._ctl_formatted = self._dialog.insert_formatted_field(
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=self._box_height,
            spin_button=True,
            value=3,
            border=border_kind,
        )
        self._ctl_formatted.text_color = StandardColor.GOLD_DARK2
        self._set_tab_index(self._ctl_formatted)

        self._ctl_formatted.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_formatted.add_event_down(self._fn_on_down)
        self._ctl_formatted.add_event_up(self._fn_on_up)
        sz = self._ctl_formatted.view.getPosSize()
        # Inserts radio buttons into dialog.

        # Inserting more then a single group of radio buttons into a dialog can be a bit buggy.
        # This in part has to do with how LibreOffice handles the Tab Indexes for the radio controls.

        # The easies solution is to set the group box tab index right before adding a set of radio controls.
        # This way between radio control set a different control is assigned a tab index before the next
        # set of radio controls is created and added.

        # When using tab indexes the tab indexes must be contiguous for each group but the next group must not
        # pick up from the same tab index that the last group finished on.

        # Alternatively it seems radio control groups can also be added with out setting a tab index as well.
        # In order for this to work a different control such as a group box must be added to the dialog between
        # adding radio control groups.
        self._ctl_gb1 = self._dialog.insert_group_box(
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=round((self._width // 2) - ((self._padding * 2) * 0.75)),
            height=100,
            label="Group Box One",
        )
        if fd is not None:
            self._ctl_gb1.set_font_descriptor(fd)
        self._set_tab_index(self._ctl_gb1)

        # insert radio buttons into group box one
        sz = self._ctl_gb1.view.getPosSize()
        self._rb1 = self._dialog.insert_radio_button(
            label="Radio Button 1",
            x=sz.X + self._padding,
            y=sz.Y + 10,
            width=sz.Width - (self._padding * 2),
            height=20,
        )
        if fd is not None:
            self._rb1.set_font_descriptor(fd)
            self._rb1.font_descriptor.underline = FontUnderlineEnum.SINGLE
            self._rb1.font_descriptor.weight = 150
        self._set_tab_index(self._rb1)
        self._rb1.state = StateKind.CHECKED
        self._rb1.add_event_item_state_changed(self._fn_on_item_changed)
        self._rb1.add_event_property_change("State", self._fn_on_property_changed)
        rb_sz = self._rb1.view.getPosSize()
        for i in range(1, 4):
            radio_btn = self._dialog.insert_radio_button(
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )
            radio_btn.set_font_descriptor(self._rb1.font_descriptor.component)
            radio_btn.text_color = StandardColor.get_random_color()
            self._set_tab_index(radio_btn)
            radio_btn.add_event_item_state_changed(self._fn_on_item_changed)
            radio_btn.add_event_property_change("State", self._fn_on_property_changed)

        sz = self._ctl_gb1.view.getPosSize()
        self._ctl_gb2 = self._dialog.insert_group_box(
            x=sz.X + sz.Width + self._padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            label="Group Box Two",
        )
        if fd is not None:
            self._ctl_gb2.set_font_descriptor(fd)
        # simplest way to break contiguous tab order and have it work
        self._set_tab_index(self._ctl_gb2)

        # insert radio buttons into group box two
        sz = self._ctl_gb2.view.getPosSize()
        self._rb2 = self._dialog.insert_radio_button(
            label="Radio Button 1",
            x=sz.X + self._padding,
            y=sz.Y + 10,
            width=sz.Width - (self._padding * 2),
            height=20,
        )
        if fd is not None:
            self._rb2.set_font_descriptor(fd)

        self._set_tab_index(self._rb2)
        self._rb2.state = StateKind.CHECKED
        self._rb2.add_event_item_state_changed(self._fn_on_item_changed)
        self._rb2.add_event_property_change("State", self._fn_on_property_changed)
        rb_sz = self._rb2.view.getPosSize()
        weight = 120
        for i in range(1, 4):
            radio_btn = self._dialog.insert_radio_button(
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )
            radio_btn.set_font_descriptor(self._rb2.font_descriptor.component)
            radio_btn.text_color = StandardColor.get_random_color()
            self._set_tab_index(radio_btn)
            radio_btn.add_event_item_state_changed(self._fn_on_item_changed)
            radio_btn.add_event_property_change("State", self._fn_on_property_changed)
            radio_btn.font_descriptor.weight = weight
            weight += 15

        sz = self._ctl_gb1.view.getPosSize()
        self._ctl_link = self._dialog.insert_hyperlink(
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=20,
            label="OOO Development Tools",
            url="https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html",
            # Enabled=False,
        )
        self._set_tab_index(self._ctl_link)
        self._ctl_link.add_event_action_performed(self._fn_on_action_general)

        sz = self._ctl_link.view.getPosSize()
        self._ctl_time = self._dialog.insert_time_field(
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=sz.Width,
            height=self._box_height,
            border=border_kind,
            time_value=datetime.datetime.now().time(),
            time_format=TimeFormatKind.LONG_12H,
        )
        if fd is not None:
            self._ctl_time.set_font_descriptor(fd)

        self._set_tab_index(self._ctl_time)

        self._ctl_time.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_time.add_event_down(self._fn_on_down)
        self._ctl_time.add_event_up(self._fn_on_up)

        sz = self._ctl_gb2.view.getPosSize()
        # file:///workspace/ooouno-dev-tools/tests/fixtures/image/img_brick.png
        pth = Path(__file__).parent.parent.parent / "fixtures" / "image" / "img_brick.png"
        self._ctl_img = self._dialog.insert_image_control(
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=120,
            height=120,
            # image_url=FileIO.fnm_to_url(pth),
            scale=ImageScaleModeEnum.ANISOTROPIC,
            border=border_kind,
        )
        # self._ctl_img.picture = FileIO.fnm_to_url(pth)
        self._ctl_img.picture = pth
        self._set_tab_index(self._ctl_img)

        self._ctl_img.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_img.add_event_mouse_exited(self._fn_on_mouse_exit)

        sz = self._ctl_img.view.getPosSize()
        self._ctl_list_box = self._dialog.insert_list_box(
            x=sz.X + sz.Width + self._padding,
            y=sz.Y,
            width=sz.Width,
            entries=["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
            drop_down=False,
            border=border_kind,
        )
        if fd is not None:
            self._ctl_list_box.set_font_descriptor(fd)
            self._ctl_list_box.font_descriptor.strikeout = FontStrikeoutEnum.SINGLE
        self._set_tab_index(self._ctl_list_box)

        sz = self._ctl_list_box.view.getPosSize()
        y = sz.Y + sz.Height + self._padding

        self._ctl_spin_btn = self._dialog.insert_spin_button(
            x=self._padding,
            y=y,
            width=30,
            height=15,
            border=BorderKind.NONE,
            orientation=OrientationKind.VERTICAL,
            min_value=self._ctl_progress.model.ProgressValueMin,
            max_value=self._ctl_progress.model.ProgressValueMax,
        )

        self._ctl_spin_btn.repeat = True
        self._ctl_spin_btn.spin_value = self._ctl_progress.value
        self._ctl_spin_btn.symbol_color = StandardColor.BLUE_LIGHT2
        self._ctl_spin_btn.add_event_adjustment_value_changed(self._fn_on_spin_adjustment)

        self._ctl_list_box.add_event_action_performed(self._fn_on_action_general)
        self._ctl_list_box.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_list_box.add_event_mouse_exited(self._fn_on_mouse_exit)
        self._ctl_list_box.add_event_item_state_changed(self._fn_on_item_changed)

        self._dialog.add_event_window_moved(self._fn_on_window_moved)
        self._dialog.add_event_window_closed(self._fn_on_window_closed)

        self._ctl_button_ok.add_event_properties_change(names=["Enabled"], cb=self._fn_on_button_properties_changed)

        self._ctl_button_ok.add_event_property_change("Enabled", self._fn_on_property_changed)
        # The vetoable event is not firing. I suspect that the Button Enable property is not a vetoable property.
        self._ctl_button_ok.add_event_vetoable_change("Enabled", self._fn_on_button_veto_property_changed)
        self._ctl_button_ok.enabled = False
        # self._ctl_button_ok.remove_event_property_change("Enabled")
        # self._ctl_button_ok.remove_event_properties_listener()
        self._ctl_button_ok.enabled = True
        self._ctl_lbl.font_descriptor.height = 15
        self._ctl_lbl.font_descriptor.slant = FontSlant.ITALIC
        self._ctl_lbl.font_descriptor.strikeout = FontStrikeoutEnum.SINGLE
        self._ctl_lbl.background_color = StandardColor.RED
        self._ctl_lbl.text_color = StandardColor.YELLOW_LIGHT2

        self._ctl_btn_cancel.font_descriptor.height = 10
        self._ctl_btn_cancel.text_color = StandardColor.RED_DARK2
        self._ctl_btn_cancel.font_descriptor.slant = FontSlant.ITALIC
        self._ctl_btn_cancel.align = AlignKind.RIGHT
        self._ctl_button_ok.align = AlignKind.LEFT

        # self._ctl_button_ok.add_event_properties_change

    def show(self) -> str:
        # Dialogs.create_dialog_peer(self._dialog)
        # window = mLo.Lo.get_frame().getContainerWindow()
        self._doc.activate()
        window = self._doc.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.set_title(self._title)
        self._dialog.set_pos_size(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.set_visible(True)
        ret = self._txt_input.text if self._dialog.execute() else ""  # type: ignore
        self._dialog.dispose()
        return ret

    # region Event Handlers
    def _init_handlers(self) -> None:
        self._fn_on_check_box_state = self.on_check_box_state
        self._fn_on_action_ok = self.on_action_ok
        self._fn_on_action_cancel = self.on_action_cancel
        self._fn_on_action_general = self.on_action_general
        self._fn_on_mouse_entered = self.on_mouse_entered
        self._fn_on_mouse_exit = self.on_mouse_exit
        self._fn_on_text_changed = self.on_text_changed
        self._fn_on_item_changed = self.on_item_changed
        self._fn_on_up = self.on_up
        self._fn_on_down = self.on_down
        self._fn_on_scroll_adjustment = self.on_scroll_adjustment
        self._fn_on_window_moved = self.on_window_moved
        self._fn_on_window_closed = self.on_window_closed
        self._fn_on_property_changed = self.on_property_changed
        self._fn_on_button_veto_property_changed = self.on_button_veto_property_changed
        self._fn_on_button_properties_changed = self.on_button_properties_changed
        self._fn_on_spin_adjustment = self.on_spin_adjustment

    def on_check_box_state(self, src: Any, event: EventArgs, control_src: CtlCheckBox, *args, **kwargs) -> None:
        itm_event = cast("ItemEvent", event.event_data)
        print("Selected:", itm_event.Selected)
        print("Source state:", control_src.state)
        # control_src.visible = False
        print("Control Name:", control_src.name)

    def on_action_ok(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        # print(event)
        print("OK:", control_src.name)

    def on_action_general(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Action:", control_src.name)

    def on_action_cancel(self, src: Any, event: EventArgs, control_src: CtlButton, *args, **kwargs) -> None:
        # print(kwargs)
        # print(src)
        # print(event.event_source)
        # print(event.source)
        # print(type(event.event_data))
        # event.event_data is com.sun.star.awt.ActionEvent
        # pylint: disable=broad-except
        try:
            print("Cancel:", control_src.name)
            print("control_src", control_src)
            print(event.event_data)
            print(event.event_data.value.ActionCommand)
        except Exception as e:
            print(e)
        # self._dialog.dispose()

    def on_mouse_entered(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        # print(control_src)
        print("Mouse Entered:", control_src.name)

    def on_mouse_exit(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        # print(control_src)
        print("Mouse Exited:", control_src.name)
        if isinstance(control_src, CtlDateField):
            print("Date Mouse Exited")

    def on_text_changed(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Text Changed:", control_src.name)
        if hasattr(control_src, "text"):
            print("Text Value:", control_src.text)
        if hasattr(control_src, "value"):
            print("Value:", control_src.value)

    def on_item_changed(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Item Changed:", control_src.name)
        itm_event = cast("ItemEvent", event.event_data)
        print("Selected:", itm_event.Selected)
        print("Highlighted:", itm_event.Highlighted)
        print("ItemId:", itm_event.ItemId)

    def on_up(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Up:", control_src.name)
        print("Value:", control_src.value)

    def on_down(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Down:", control_src.name)
        print("Value:", control_src.value)

    def on_scroll_adjustment(self, src: Any, event: EventArgs, control_src: CtlScrollBar, *args, **kwargs) -> None:
        # print("Scroll:", control_src.name)
        a_event = cast("AdjustmentEvent", event.event_data)
        self._ctl_spin_btn.spin_value = a_event.Value
        self._ctl_progress.value = a_event.Value

    def on_spin_adjustment(self, src: Any, event: EventArgs, control_src: CtlScrollBar, *args, **kwargs) -> None:
        # print("Scroll:", control_src.name)
        a_event = cast("AdjustmentEvent", event.event_data)
        self._ctl_scroll_progress.value = a_event.Value
        self._ctl_progress.value = a_event.Value

    def on_window_moved(self, src: Any, event: EventArgs, control_src: CtlDialog, *args, **kwargs) -> None:
        # print("Scroll:", control_src.name)
        a_event = cast("WindowEvent", event.event_data)
        print("Window Height:", a_event.Height)
        print("Window Width:", a_event.Width)
        print("Window X:", a_event.X)
        print("Window Y:", a_event.Y)

    def on_window_closed(self, src: Any, event: EventArgs, control_src: CtlDialog, *args, **kwargs) -> None:
        # a_event = cast("WindowEvent", event.event_data)
        # this event seems to get called if the application (calc) looses focus. Not sure why.

        print("Window Closed")

    def on_property_changed(
        self, src: Any, event: EventArgs, property_name: str, control_src: Any, component: Any, *args, **kwargs
    ) -> None:
        itm_event = cast("PropertyChangeEvent", event.event_data)
        print(
            "Property Changed:",
            property_name,
            ",",
            " New Value:",
            itm_event.NewValue,
            ",",
            "Old Value:",
            itm_event.OldValue,
        )

    def on_button_properties_changed(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        itm_event = cast("Tuple[PropertyChangeEvent]", event.event_data)
        for itm in itm_event:
            print(
                "Button Properties Changed:",
                itm.PropertyName,
                ",",
                " New Value:",
                itm.NewValue,
                ",",
                "Old Value:",
                itm.OldValue,
            )

    def on_button_veto_property_changed(
        self, src: Any, event: EventArgs, property_name: str, control_src: Any, component: Any, *args, **kwargs
    ) -> None:
        itm_event = cast("PropertyChangeEvent", event.event_data)
        print(
            "Button Veto Property Changed:",
            property_name,
            ",",
            "New Value:",
            itm_event.NewValue,
            ",",
            "Old Value:",
            itm_event.OldValue,
        )

    # endregion Event Handlers

    def _set_tab_index(self, ctl: DialogControlBase) -> None:
        ctl.tab_index = self._tab_index
        self._tab_index += 1


def main():
    # It is important to activate windows before displaying the dialog when using multi-documents.
    # In this case the windows are activated in the show method.
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = CalcDoc.create_doc(visible=True)

        # inst2 = mLo.Lo.create_lo_instance()
        # doc2 = CalcDoc.create_doc(lo_inst=inst2, visible=True)
        # run(doc2)

        # inst3 = mLo.Lo.create_lo_instance()
        # doc3 = WriteDoc.create_doc(lo_inst=inst3, visible=True)
        # run(doc3)
        run(doc)


def run(doc: WriteDoc | CalcDoc) -> None:
    inst = Runner(doc, "title", "msg", "input_value")
    print(inst.show())


if __name__ == "__main__":
    main()
