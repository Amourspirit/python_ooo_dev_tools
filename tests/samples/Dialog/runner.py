from __future__ import annotations
import datetime
from typing import Any, TYPE_CHECKING, cast, Tuple
from pathlib import Path
import uno  # pylint: disable=unused-import

from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.push_button_type import PushButtonType

from ooodev.dialog import Dialogs, ImageScaleModeEnum, BorderKind, DateFormatKind, TimeFormatKind, StateKind
from ooodev.utils import lo as mLo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.utils.file_io import FileIO
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.color import StandardColor
from ooodev.dialog.dl_control.ctl_date_field import CtlDateField


if TYPE_CHECKING:
    from com.sun.star.awt import ItemEvent
    from com.sun.star.awt import AdjustmentEvent
    from com.sun.star.awt import WindowEvent
    from com.sun.star.beans import PropertyChangeEvent
    from ooodev.dialog.dl_control.ctl_button import CtlButton
    from ooodev.dialog.dl_control.ctl_check_box import CtlCheckBox
    from ooodev.dialog.dl_control.ctl_combo_box import CtlComboBox
    from ooodev.dialog.dl_control.ctl_scroll_bar import CtlScrollBar
    from ooodev.dialog.dl_control.ctl_dialog import CtlDialog
    from ooodev.dialog.dl_control.ctl_base import DialogControlBase


class Runner:
    # pylint: disable=unused-argument
    def __init__(
        self,
        title: str,
        msg: str,
        input_value: str = "",
        ok_lbl: str = "OK",
        cancel_lbl: str = "Cancel",
        is_password: bool = False,
    ) -> None:
        self._init_handlers()

        # self._dialog = cast(
        #     "UnoControlDialog",
        #     mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        # )
        # dialog_model = mLo.Lo.create_instance_mcf(
        #     XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
        # )

        # self._dialog.setModel(dialog_model)
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
        self._dialog = Dialogs.create_dialog(
            x=-1,
            y=-1,
            width=self._width,
            height=self._height,
            title=self._title,
        )

        self._ctl_lbl = Dialogs.insert_label(
            dialog_ctrl=self._dialog.control,
            label=msg,
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=20,
        )
        self._set_tab_index(self._ctl_lbl)
        self._ctl_lbl.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_lbl.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_lbl.view.getPosSize()
        if is_password:
            self._txt_input = Dialogs.insert_password_field(
                dialog_ctrl=self._dialog.control,
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
            )
        else:
            self._txt_input = Dialogs.insert_text_field(
                dialog_ctrl=self._dialog.control,
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
            )
        self._set_tab_index(self._txt_input)
        self._txt_input.add_event_text_changed(self._fn_on_text_changed)

        self._ctl_btn_cancel = Dialogs.insert_button(
            dialog_ctrl=self._dialog.control,
            label=cancel_lbl,
            x=self._width - self._btn_width - self._margin,
            y=self._height - self._btn_height - self._margin,
            width=self._btn_width,
            height=self._btn_height,
            # btn_type=PushButtonType.CANCEL,
        )
        self._set_tab_index(self._ctl_btn_cancel)
        self._ctl_btn_cancel.view.setActionCommand("Cancel")
        self._ctl_btn_cancel.add_event_action_performed(self._fn_on_action_cancel)
        self._ctl_btn_cancel.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_btn_cancel.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_btn_cancel.view.getPosSize()
        self._ctl_button_ok = Dialogs.insert_button(
            dialog_ctrl=self._dialog.control,
            label=ok_lbl,
            x=sz.X - sz.Width - self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )
        self._set_tab_index(self._ctl_button_ok)
        self._ctl_button_ok.add_event_action_performed(self._fn_on_action_ok)
        self._ctl_button_ok.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_button_ok.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._txt_input.view.getPosSize()
        # ctl_button_ok.width += 30
        # ctl_button_ok.x -= 30
        self._ctl_chk1 = Dialogs.insert_check_box(
            dialog_ctrl=self._dialog.control,
            label="Check Box 1",
            x=sz.X,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=20,
            tri_state=False,
            state=Dialogs.StateEnum.CHECKED,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_chk1)

        sz = self._ctl_chk1.view.getPosSize()
        self._ctl_chk2 = Dialogs.insert_check_box(
            dialog_ctrl=self._dialog.control,
            label="Check Box 2",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=False,
            state=Dialogs.StateEnum.NOT_CHECKED,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_chk2)

        sz = self._ctl_chk2.view.getPosSize()
        self._ctl_chk3 = Dialogs.insert_check_box(
            dialog_ctrl=self._dialog.control,
            label="Check Box 3",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=True,
            state=Dialogs.StateEnum.DONT_KNOW,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_chk3)
        self._ctl_chk1.add_event_item_state_changed(self._fn_on_check_box_state)
        self._ctl_chk2.add_event_item_state_changed(self._fn_on_check_box_state)
        self._ctl_chk3.add_event_item_state_changed(self._fn_on_check_box_state)

        sz = self._ctl_chk1.view.getPosSize()
        self._ctl_date = Dialogs.insert_date_field(
            dialog_ctrl=self._dialog.control,
            x=sz.Width + self._padding,
            y=sz.Y,
            width=190,
            height=self._box_height,
            date_value=datetime.datetime.now(),
            border=border_kind,
        )
        self._set_tab_index(self._ctl_date)
        self._ctl_date.date_format = DateFormatKind.DIN_5008_YY_MM_DD
        dt = datetime.datetime.now()
        self._ctl_date.date = datetime.datetime(dt.year - 1, dt.month, dt.day)

        self._ctl_date.add_event_down(self._fn_on_down)
        self._ctl_date.add_event_up(self._fn_on_up)
        self._ctl_date.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_date.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_date.view.getPosSize()
        self._ctl_currency = Dialogs.insert_currency_field(
            dialog_ctrl=self._dialog.control,
            x=sz.Width + sz.X + self._padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            value=123.45,
            spin_button=True,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_currency)
        sz = self._ctl_currency.view.getPosSize()
        self._ctl_currency.add_event_down(self._fn_on_down)
        self._ctl_currency.add_event_up(self._fn_on_up)
        self._ctl_pattern = Dialogs.insert_pattern_field(
            dialog_ctrl=self._dialog.control,
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=sz.Width,
            height=sz.Height,
            edit_mask="NNLNNLLLLL",
            literal_mask="__.__.2025",
            border=border_kind,
        )
        self._set_tab_index(self._ctl_pattern)
        self._ctl_pattern.add_event_down(self._fn_on_down)
        self._ctl_pattern.add_event_up(self._fn_on_up)
        self._ctl_pattern.add_event_text_changed(self._fn_on_text_changed)

        sz_date = self._ctl_date.view.getPosSize()
        self._ctl_num_field = Dialogs.insert_numeric_field(
            dialog_ctrl=self._dialog.control,
            x=sz_date.X,
            y=sz_date.Y + sz_date.Height + self._padding,
            width=sz.Width,
            height=self._box_height,
            value=123,
            spin_button=True,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_num_field)
        self._ctl_num_field.add_event_down(self._fn_on_down)
        self._ctl_num_field.add_event_up(self._fn_on_up)
        self._ctl_num_field.add_event_text_changed(self._fn_on_text_changed)

        sz_numeric = self._ctl_num_field.view.getPosSize()
        # sz_fmt = ctl_formatted.getPosSize()
        self._ctl_combo1 = Dialogs.insert_combo_box(
            dialog_ctrl=self._dialog.control,
            x=self._margin,
            y=sz_numeric.Height + sz_numeric.Y + self._padding,
            width=200,
            height=self._box_height,
            entries=["Item 1", "Item 2", "Item 3"],
            border=border_kind,
        )
        self._set_tab_index(self._ctl_combo1)
        self._ctl_combo1.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_combo1.add_event_item_state_changed(self._fn_on_item_changed)

        sz = self._ctl_combo1.view.getPosSize()

        self._ctl_progress = Dialogs.insert_progress_bar(
            dialog_ctrl=self._dialog.control,
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

        self._ctl_scroll_progress = Dialogs.insert_scroll_bar(
            dialog_ctrl=self._dialog.control,
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

        self._ctl_file = Dialogs.insert_file_control(
            dialog_ctrl=self._dialog.control,
            x=sz.X,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=self._box_height,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_file)
        self._ctl_file.text = "file:///workspace/ooouno-dev-tools/tests/fixtures/image/img_brick.png"
        self._ctl_file.add_event_text_changed(self._fn_on_text_changed)
        sz = self._ctl_file.view.getPosSize()
        self._ctl_ln = Dialogs.insert_fixed_line(
            dialog_ctrl=self._dialog.control,
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=self._width - (self._margin * 2),
            height=1,
        )
        self._set_tab_index(self._ctl_ln)

        sz = self._ctl_ln.view.getPosSize()
        self._ctl_formatted = Dialogs.insert_formatted_field(
            dialog_ctrl=self._dialog.control,
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=self._box_height,
            spin_button=True,
            value=3,
            border=border_kind,
        )
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
        self._ctl_gb1 = Dialogs.insert_group_box(
            dialog_ctrl=self._dialog.control,
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=round((self._width // 2) - ((self._padding * 2) * 0.75)),
            height=100,
            label="Group Box One",
        )
        self._set_tab_index(self._ctl_gb1)

        # insert radio buttons into group box one
        sz = self._ctl_gb1.view.getPosSize()
        self._rb1 = Dialogs.insert_radio_button(
            dialog_ctrl=self._dialog.control,
            label="Radio Button 1",
            x=sz.X + self._padding,
            y=sz.Y + 10,
            width=sz.Width - (self._padding * 2),
            height=20,
        )
        self._set_tab_index(self._rb1)
        self._rb1.state = StateKind.CHECKED
        self._rb1.add_event_item_state_changed(self._fn_on_item_changed)
        rb_sz = self._rb1.view.getPosSize()
        for i in range(1, 4):
            radio_btn = Dialogs.insert_radio_button(
                dialog_ctrl=self._dialog.control,
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )
            self._set_tab_index(radio_btn)
            radio_btn.add_event_item_state_changed(self._fn_on_item_changed)

        sz = self._ctl_gb1.view.getPosSize()
        self._ctl_gb2 = Dialogs.insert_group_box(
            dialog_ctrl=self._dialog.control,
            x=sz.X + sz.Width + self._padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            label="Group Box Two",
        )
        # simplest way to break contiguous tab order and have it work
        self._set_tab_index(self._ctl_gb2)

        # insert radio buttons into group box two
        sz = self._ctl_gb2.view.getPosSize()
        self._rb2 = Dialogs.insert_radio_button(
            dialog_ctrl=self._dialog.control,
            label="Radio Button 1",
            x=sz.X + self._padding,
            y=sz.Y + 10,
            width=sz.Width - (self._padding * 2),
            height=20,
        )

        self._set_tab_index(self._rb2)
        self._rb2.state = StateKind.CHECKED
        self._rb2.add_event_item_state_changed(self._fn_on_item_changed)
        rb_sz = self._rb2.view.getPosSize()
        for i in range(1, 4):
            radio_btn = Dialogs.insert_radio_button(
                dialog_ctrl=self._dialog.control,
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )
            self._set_tab_index(radio_btn)
            radio_btn.add_event_item_state_changed(self._fn_on_item_changed)

        sz = self._ctl_gb1.view.getPosSize()
        self._ctl_link = Dialogs.insert_hyperlink(
            dialog_ctrl=self._dialog.control,
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
        self._ctl_time = Dialogs.insert_time_field(
            dialog_ctrl=self._dialog.control,
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=sz.Width,
            height=self._box_height,
            border=border_kind,
            time_value=datetime.datetime.now().time(),
            time_format=TimeFormatKind.LONG_12H,
        )
        self._set_tab_index(self._ctl_time)

        self._ctl_time.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_time.add_event_down(self._fn_on_down)
        self._ctl_time.add_event_up(self._fn_on_up)

        sz = self._ctl_gb2.view.getPosSize()
        # file:///workspace/ooouno-dev-tools/tests/fixtures/image/img_brick.png
        pth = Path(__file__).parent.parent.parent / "fixtures" / "image" / "img_brick.png"
        self._ctl_img = Dialogs.insert_image_control(
            dialog_ctrl=self._dialog.control,
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
        self._ctl_list_box = Dialogs.insert_list_box(
            dialog_ctrl=self._dialog.control,
            x=sz.X + sz.Width + self._padding,
            y=sz.Y,
            width=sz.Width,
            entries=["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
            drop_down=False,
            border=border_kind,
        )
        self._set_tab_index(self._ctl_list_box)

        self._ctl_list_box.add_event_action_performed(self._fn_on_action_general)
        self._ctl_list_box.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_list_box.add_event_mouse_exited(self._fn_on_mouse_exit)
        self._ctl_list_box.add_event_item_state_changed(self._fn_on_item_changed)

        self._dialog.add_event_window_moved(self._fn_on_window_moved)
        self._dialog.add_event_window_closed(self._fn_on_window_closed)

        self._ctl_button_ok.add_event_properties_change(names=["Enabled"], cb=self._fn_on_button_properties_changed)

        self._ctl_button_ok.add_event_property_change("Enabled", self._fn_on_button_property_changed)
        # The vetoable event is not firing. I suspect that the Button Enable property is not a vetoable property.
        self._ctl_button_ok.add_event_vetoable_change("Enabled", self._fn_on_button_veto_property_changed)
        self._ctl_button_ok.enabled = False
        # self._ctl_button_ok.remove_event_property_change("Enabled")
        # self._ctl_button_ok.remove_event_properties_listener()
        self._ctl_button_ok.enabled = True

        # self._ctl_button_ok.add_event_properties_change

    def show(self) -> str:
        # Dialogs.create_dialog_peer(self._dialog)
        window = mLo.Lo.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.control.setTitle(self._title)
        self._dialog.control.setPosSize(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.control.setVisible(True)
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
        self._fn_on_button_property_changed = self.on_button_property_changed
        self._fn_on_button_veto_property_changed = self.on_button_veto_property_changed
        self._fn_on_button_properties_changed = self.on_button_properties_changed

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
        self._ctl_progress.value = a_event.Value

    def on_window_moved(self, src: Any, event: EventArgs, control_src: CtlDialog, *args, **kwargs) -> None:
        # print("Scroll:", control_src.name)
        a_event = cast("WindowEvent", event.event_data)
        print("Window Height:", a_event.Height)
        print("Window Height:", a_event.Width)
        print("Window X:", a_event.X)
        print("Window Y:", a_event.Y)

    def on_window_closed(self, src: Any, event: EventArgs, control_src: CtlDialog, *args, **kwargs) -> None:
        # a_event = cast("WindowEvent", event.event_data)
        # this event seems to get called if the application (calc) looses focus. Not sure why.

        print("Window Closed")

    def on_button_property_changed(
        self, src: Any, event: EventArgs, property_name: str, control_src: Any, component: Any, *args, **kwargs
    ) -> None:
        itm_event = cast("PropertyChangeEvent", event.event_data)
        print(
            "Button Property Changed:",
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
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = Calc.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        run()


def run() -> None:
    inst = Runner("title", "msg", "input_value")
    print(inst.show())


if __name__ == "__main__":
    main()
