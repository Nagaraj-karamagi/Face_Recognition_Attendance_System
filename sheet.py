import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, BOTH, RIGHT, Y, X, BOTTOM


class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Sheet Viewer")
        self.root.geometry("900x600")

        # Get project base directory
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_file = os.path.join(self.base_dir, "attendance.xlsx")

        self.setup_ui()

    # ---------------- UI SETUP ---------------- #

    def setup_ui(self):
        # Text widget
        self.text_widget = tk.Text(
            self.root,
            wrap="none",
            font=("Courier New", 11),
            padx=10,
            pady=10
        )
        self.text_widget.pack(fill=BOTH, expand=True)

        # Vertical Scrollbar
        v_scroll = tk.Scrollbar(self.root, command=self.text_widget.yview)
        v_scroll.pack(side=RIGHT, fill=Y)
        self.text_widget.configure(yscrollcommand=v_scroll.set)

        # Horizontal Scrollbar
        h_scroll = tk.Scrollbar(self.root, orient="horizontal", command=self.text_widget.xview)
        h_scroll.pack(side=BOTTOM, fill=X)
        self.text_widget.configure(xscrollcommand=h_scroll.set)

        # Load and display data
        self.load_and_display()

    # ---------------- LOAD EXCEL ---------------- #

    def load_and_display(self):
        if not os.path.exists(self.excel_file):
            messagebox.showerror("Error", "attendance.xlsx not found!")
            self.text_widget.insert("1.0", "Attendance file not found.")
            return

        try:
            df = pd.read_excel(self.excel_file)
            if df.empty:
                self.text_widget.insert("1.0", "No attendance records found.")
            else:
                data_str = df.to_string(index=False)
                self.text_widget.insert("1.0", data_str)
        except Exception as e:
            messagebox.showerror("Error", str(e))


# ---------------- MAIN ---------------- #

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
