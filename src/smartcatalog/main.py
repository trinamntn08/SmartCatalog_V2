import tkinter as tk
from smartcatalog.ui.main_window import create_main_window

def start_ui():
    root = tk.Tk()
    root.title("SmartCatalog – Trích xuất & Đối chiếu sản phẩm")
    create_main_window(root)
    root.mainloop()

if __name__ == "__main__":
    start_ui()
