from tkinter import *


class Window:
    """Creates a dialog Window"""
    def __init__(self, title: str = "Input", input_msg: str = "Enter Input:", is_password: bool = False) -> None:
        """
        Class Constructor

        Args:
            title (str, optional): Title to display for Dialog. Defaults to "Input".
            input_msg (_type_, optional): Message to display for Dialog. Defaults to "Enter Input:".
            is_password (bool, optional): Determines if input box is masked for password input. Defaults to False.
        """
        self.root = Tk()
        self.root.geometry("300x70")
        self.root.title(title)
        self.parent = Frame(self.root, padx=10, pady=10)
        self.parent.pack(fill=BOTH, expand=True)
        mask = "*" if is_password else ""
        self.password = StringVar()  # Password variable
        self.passEntry = self._make_entry(self.parent, input_msg, 16, show=mask, textvariable=self.password)
        submit = Button(self.root, text="OK", command=self.get_input)

        # passEntry.pack(pady=12, padx=8)
        submit.pack()

        self.root.mainloop()

    def get_input(self) -> str:
        """
        Gets the input from the dialog

        Returns:
            str: Dialog Input or empty string.
        """
        p = self.password.get()  # get password from entry
        self.root.quit()
        # print("Password:", p)
        return p

    def _make_entry(self, parent, caption, width=None, **options):
        Label(parent, text=caption).pack(side=TOP)
        entry = Entry(parent, **options)
        if width:
            entry.config(width=width)
        entry.pack(side=TOP, padx=10, fill=BOTH)
        return entry
