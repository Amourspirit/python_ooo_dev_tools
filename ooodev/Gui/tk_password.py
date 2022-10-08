import tkinter as tk
import tkinter.simpledialog

def main() -> str | None:
    tk.Tk().withdraw()
    result = tkinter.simpledialog.askstring("Password", "Enter password:", show='*')
    return result

if __name__ == "__main__":
    print(main())