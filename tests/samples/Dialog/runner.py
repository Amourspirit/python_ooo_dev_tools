from __future__ import annotations
import datetime
from typing import Any, TYPE_CHECKING, cast
from pathlib import Path
import datetime
from ooodev.dialog import Dialogs, ImageScaleModeEnum, BorderKind, DateFormatKind
from ooodev.utils import lo as mLo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.utils.file_io import FileIO
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.color import StandardColor

from com.sun.star.awt import XControlModel
from com.sun.star.awt import XDialog

from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.push_button_type import PushButtonType

if TYPE_CHECKING:
    from ooodev.dialog.dl_control.ctl_button import CtlButton
    from ooodev.dialog.dl_control.ctl_check_box import CtlCheckBox
    from ooodev.dialog.dl_control.ctl_combo_box import CtlComboBox
    from com.sun.star.awt import UnoControlDialog
    from com.sun.star.awt import ItemEvent


class Runner:
    def __init__(
        self,
        title: str,
        msg: str,
        input_value: str = "",
        ok_lbl: str = "OK",
        cancel_lbl: str = "Cancel",
        is_password: bool = False,
    ) -> None:
        """
        Displays an input box and returns the results.

        Args:
            title (str): Title for the dialog
            msg (str): Message to display such as "Input your Name"
            input_value (str, optional): Value of input box when first displayed.
            ok_lbl (str, optional): OK button Label. Defaults to "OK".
            cancel_lbl (str, optional): Cancel Button Label. Defaults to "Cancel".
            is_password (bool, optional): Determines if the input box is masked for password input. Defaults to False.

        Returns:
            str: The value of input or empty string.
        """

        self._init_handlers()

        self._dialog = cast(
            "UnoControlDialog",
            mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        )
        dialog_model = mLo.Lo.create_instance_mcf(
            XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
        )

        self._dialog.setModel(dialog_model)
        border_kind = BorderKind.NONE
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

        self._ctl_lbl = Dialogs.insert_label(
            dialog_ctrl=self._dialog,
            label=msg,
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=20,
        )
        self._ctl_lbl.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_lbl.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_lbl.view.getPosSize()
        if is_password:
            self._txt_input = Dialogs.insert_password_field(
                dialog_ctrl=self._dialog,
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
            )
        else:
            self._txt_input = Dialogs.insert_text_field(
                dialog_ctrl=self._dialog,
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
            )

        self._txt_input.add_event_text_changed(self._fn_on_text_changed)

        self._ctl_btn_cancel = Dialogs.insert_button(
            dialog_ctrl=self._dialog,
            label=cancel_lbl,
            x=self._width - self._btn_width - self._margin,
            y=self._height - self._btn_height - self._margin,
            width=self._btn_width,
            height=self._btn_height,
            # btn_type=PushButtonType.CANCEL,
        )
        self._ctl_btn_cancel.view.setActionCommand("Cancel")
        self._ctl_btn_cancel.add_event_action_performed(self._fn_on_action_cancel)
        self._ctl_btn_cancel.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_btn_cancel.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._ctl_btn_cancel.view.getPosSize()
        self._ctl_button_ok = Dialogs.insert_button(
            dialog_ctrl=self._dialog,
            label=ok_lbl,
            x=sz.X - sz.Width - self._margin,
            y=sz.Y,
            width=self._btn_width,
            height=self._btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )
        self._ctl_button_ok.add_event_action_performed(self._fn_on_action_ok)
        self._ctl_button_ok.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_button_ok.add_event_mouse_exited(self._fn_on_mouse_exit)
        sz = self._txt_input.view.getPosSize()
        # ctl_button_ok.width += 30
        # ctl_button_ok.x -= 30
        self._ctl_chk1 = Dialogs.insert_check_box(
            dialog_ctrl=self._dialog,
            label="Check Box 1",
            x=sz.X,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=20,
            tri_state=False,
            state=Dialogs.StateEnum.CHECKED,
            border=border_kind,
        )

        sz = self._ctl_chk1.view.getPosSize()
        self._ctl_chk2 = Dialogs.insert_check_box(
            dialog_ctrl=self._dialog,
            label="Check Box 2",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=False,
            state=Dialogs.StateEnum.NOT_CHECKED,
            border=border_kind,
        )

        sz = self._ctl_chk2.view.getPosSize()
        self._ctl_chk3 = Dialogs.insert_check_box(
            dialog_ctrl=self._dialog,
            label="Check Box 3",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=True,
            state=Dialogs.StateEnum.DONT_KNOW,
            border=border_kind,
        )
        self._ctl_chk1.add_event_item_state_changed(self._fn_on_check_box_state)
        self._ctl_chk2.add_event_item_state_changed(self._fn_on_check_box_state)
        self._ctl_chk3.add_event_item_state_changed(self._fn_on_check_box_state)

        sz = self._ctl_chk1.view.getPosSize()
        self._ctl_date = Dialogs.insert_date_field(
            dialog_ctrl=self._dialog,
            x=sz.Width + self._padding,
            y=sz.Y,
            width=190,
            height=self._box_height,
            date_value=datetime.datetime.now(),
            border=border_kind,
        )
        self._ctl_date.date_format = DateFormatKind.DIN_5008_YY_MM_DD
        dt = datetime.datetime.now()
        self._ctl_date.date = datetime.datetime(dt.year - 1, dt.month, dt.day)

        self._ctl_date.add_event_down(self._fn_on_down)
        self._ctl_date.add_event_up(self._fn_on_up)
        self._ctl_date.add_event_text_changed(self._fn_on_text_changed)
        sz = self._ctl_date.view.getPosSize()
        self._ctl_currency = Dialogs.insert_currency_field(
            dialog_ctrl=self._dialog,
            x=sz.Width + sz.X + self._padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            value=123.45,
            spin_button=True,
            border=border_kind,
        )
        sz = self._ctl_currency.view.getPosSize()
        self._ctl_currency.add_event_down(self._fn_on_down)
        self._ctl_currency.add_event_up(self._fn_on_up)
        self._ctl_pattern = Dialogs.insert_pattern_field(
            dialog_ctrl=self._dialog,
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=sz.Width,
            height=sz.Height,
            edit_mask="NNLNNLLLLL",
            literal_mask="__.__.2025",
            border=border_kind,
        )

        self._ctl_pattern.add_event_down(self._fn_on_down)
        self._ctl_pattern.add_event_up(self._fn_on_up)
        self._ctl_pattern.add_event_text_changed(self._fn_on_text_changed)

        sz_date = self._ctl_date.view.getPosSize()
        self._ctl_num_field = Dialogs.insert_numeric_field(
            dialog_ctrl=self._dialog,
            x=sz_date.X,
            y=sz_date.Y + sz_date.Height + self._padding,
            width=sz.Width,
            height=self._box_height,
            value=123,
            spin_button=True,
            border=border_kind,
        )
        self._ctl_num_field.add_event_down(self._fn_on_down)
        self._ctl_num_field.add_event_up(self._fn_on_up)
        self._ctl_num_field.add_event_text_changed(self._fn_on_text_changed)

        sz_numeric = self._ctl_num_field.view.getPosSize()
        # sz_fmt = ctl_formatted.getPosSize()
        self._ctl_combo1 = Dialogs.insert_combo_box(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=sz_numeric.Height + sz_numeric.Y + self._padding,
            width=200,
            height=self._box_height,
            entries=["Item 1", "Item 2", "Item 3"],
            border=border_kind,
        )
        self._ctl_combo1.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_combo1.add_event_item_state_changed(self._fn_on_item_changed)

        sz = self._ctl_combo1.view.getPosSize()

        self._ctl_progress = Dialogs.insert_progress_bar(
            dialog_ctrl=self._dialog,
            x=sz_date.X,
            y=sz.Y,
            width=400,
            height=self._box_height,
            min=1,
            value=67,
            border=border_kind,
        )
        self._ctl_progress.fill_color = StandardColor.GREEN
        self._ctl_progress.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_progress.add_event_mouse_exited(self._fn_on_mouse_exit)

        self._ctl_file = Dialogs.insert_file_control(
            dialog_ctrl=self._dialog,
            x=sz.X,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=self._box_height,
            border=border_kind,
        )
        self._ctl_file.text = "file:///workspace/ooouno-dev-tools/tests/fixtures/image/img_brick.png"
        self._ctl_file.add_event_text_changed(self._fn_on_text_changed)
        sz = self._ctl_file.view.getPosSize()
        self._ctl_ln = Dialogs.insert_fixed_line(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=self._width - (self._margin * 2),
            height=1,
        )

        sz = self._ctl_ln.view.getPosSize()
        self._ctl_formatted = Dialogs.insert_formatted_field(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=self._box_height,
            spin_button=True,
            value=3,
            border=border_kind,
        )

        self._ctl_formatted.add_event_text_changed(self._fn_on_text_changed)
        self._ctl_formatted.add_event_down(self._fn_on_down)
        self._ctl_formatted.add_event_up(self._fn_on_up)
        sz = self._ctl_formatted.view.getPosSize()
        self._ctl_gb1 = Dialogs.insert_group_box(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=round((self._width // 2) - ((self._padding * 2) * 0.75)),
            height=100,
            label="Group Box One",
        )

        # insert radio buttons into group box one
        sz = self._ctl_gb1.view.getPosSize()
        rb1 = Dialogs.insert_radio_button(
            dialog_ctrl=self._dialog,
            label="Radio Button 1",
            x=sz.X + self._padding,
            y=sz.Y + 10,
            width=sz.Width - (self._padding * 2),
            height=20,
        )
        rb1.add_event_item_state_changed(self._fn_on_item_changed)
        rb_sz = rb1.view.getPosSize()
        for i in range(1, 4):
            radio_btn = Dialogs.insert_radio_button(
                dialog_ctrl=self._dialog,
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )
            radio_btn.add_event_item_state_changed(self._fn_on_item_changed)

        sz = self._ctl_gb1.view.getPosSize()
        self._ctl_gb2 = Dialogs.insert_group_box(
            dialog_ctrl=self._dialog,
            x=sz.X + sz.Width + self._padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            label="Group Box Two",
        )

        # insert radio buttons into group box two
        sz = self._ctl_gb2.view.getPosSize()
        self._rb2 = Dialogs.insert_radio_button(
            dialog_ctrl=self._dialog,
            label="Radio Button 1",
            x=sz.X + self._padding,
            y=sz.Y + 10,
            width=sz.Width - (self._padding * 2),
            height=20,
        )
        self._rb2.add_event_item_state_changed(self._fn_on_item_changed)
        rb_sz = self._rb2.view.getPosSize()
        for i in range(1, 4):
            radio_btn = Dialogs.insert_radio_button(
                dialog_ctrl=self._dialog,
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )
            radio_btn.add_event_item_state_changed(self._fn_on_item_changed)

        sz = self._ctl_gb1.view.getPosSize()
        self._ctl_link = Dialogs.insert_hyperlink(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=sz.Height + sz.Y + self._padding,
            width=200,
            height=20,
            label="OOO Development Tools",
            url="https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html",
            # Enabled=False,
        )
        self._ctl_link.add_event_action_performed(self._fn_on_action_general)
        sz = self._ctl_gb2.view.getPosSize()
        # file:///workspace/ooouno-dev-tools/tests/fixtures/image/img_brick.png
        pth = Path(__file__).parent.parent.parent / "fixtures" / "image" / "img_brick.png"
        self._ctl_img = Dialogs.insert_image_control(
            dialog_ctrl=self._dialog,
            x=sz.X,
            y=sz.Y + sz.Height + self._padding,
            width=120,
            height=120,
            image_url=FileIO.fnm_to_url(pth),
            scale=ImageScaleModeEnum.ANISOTROPIC,
            border=border_kind,
        )
        self._ctl_img.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_img.add_event_mouse_exited(self._fn_on_mouse_exit)

        sz = self._ctl_img.view.getPosSize()
        self._ctl_list_box = Dialogs.insert_list_box(
            dialog_ctrl=self._dialog,
            x=sz.X + sz.Width + self._padding,
            y=sz.Y,
            width=sz.Width,
            entries=["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
            drop_down=False,
            border=border_kind,
        )

        self._ctl_list_box.add_event_action_performed(self._fn_on_action_general)
        self._ctl_list_box.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._ctl_list_box.add_event_mouse_exited(self._fn_on_mouse_exit)
        self._ctl_list_box.add_event_item_state_changed(self._fn_on_item_changed)

    def show(self) -> str:
        window = mLo.Lo.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.setTitle(self._title)
        self._dialog.setPosSize(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.setVisible(True)
        ret = self._txt_input.getModel().Text if self._dialog.execute() else ""  # type: ignore
        self._dialog.dispose()
        return ret

    # region Event Handlers
    def _init_handlers(self) -> None:
        def _on_check_box_state(src: Any, event: EventArgs, control_src: CtlCheckBox, *args, **kwargs):
            self.on_check_box_state(src, event, control_src, *args, **kwargs)

        def _on_action_ok(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_action_ok(src, event, control_src, *args, **kwargs)

        def _on_action_cancel(src: Any, event: EventArgs, control_src: CtlButton, *args, **kwargs) -> None:
            self.on_action_cancel(src, event, control_src, *args, **kwargs)

        def _on_action_general(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_action_general(src, event, control_src, *args, **kwargs)

        def _on_mouse_entered(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_mouse_entered(src, event, control_src, *args, **kwargs)

        def _on_mouse_exit(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_mouse_exit(src, event, control_src, *args, **kwargs)

        def _on_text_changed(src: Any, event: EventArgs, control_src: CtlComboBox, *args, **kwargs) -> None:
            self.on_text_changed(src, event, control_src, *args, **kwargs)

        def _on_item_changed(src: Any, event: EventArgs, control_src: CtlComboBox, *args, **kwargs) -> None:
            self.on_item_changed(src, event, control_src, *args, **kwargs)

        def _on_down(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_down(src, event, control_src, *args, **kwargs)

        def _on_up(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_up(src, event, control_src, *args, **kwargs)

        self._fn_on_check_box_state = _on_check_box_state
        self._fn_on_action_ok = _on_action_ok
        self._fn_on_action_cancel = _on_action_cancel
        self._fn_on_action_general = _on_action_general
        self._fn_on_mouse_entered = _on_mouse_entered
        self._fn_on_mouse_exit = _on_mouse_exit
        self._fn_on_text_changed = _on_text_changed
        self._fn_on_item_changed = _on_item_changed
        self._fn_on_up = _on_up
        self._fn_on_down = _on_down

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

    # endregion Event Handlers


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
