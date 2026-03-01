import tkinter as tk
from gui import MarksApp


def main():
    root = tk.Tk()
    app = MarksApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()