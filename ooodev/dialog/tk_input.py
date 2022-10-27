from ooodev.mock import mock_g
from ..mock import mock_g

try:
    import tkinter as tk
except ImportError:
    if mock_g.DOCS_BUILDING:
        # ignoring simply because windows and sphinx don't see tkinter when python interpreter is pointed to LibreOffice python.exe
        pass
    else:
        raise


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
        self.root = tk.Tk()
        self.root.geometry("300x70")
        self.root.title(title)
        self.parent = tk.Frame(self.root, padx=10, pady=10)
        self.parent.pack(fill=tk.BOTH, expand=True)
        mask = "*" if is_password else ""
        self.password = tk.StringVar()  # Password variable
        self.passEntry = self._make_entry(self.parent, input_msg, 16, show=mask, textvariable=self.password)
        submit = tk.Button(self.root, text="OK", command=self.get_input)

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
        tk.Label(parent, text=caption).pack(side=tk.TOP)
        entry = tk.Entry(parent, **options)
        if width:
            entry.config(width=width)
        entry.pack(side=tk.TOP, padx=10, fill=tk.BOTH)
        return entry
