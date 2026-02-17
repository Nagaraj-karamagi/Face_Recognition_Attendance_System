import os
import sys
import subprocess
from tkinter import Tk, Label, Button, Frame, messagebox
from PIL import Image, ImageTk


class FaceRecognitionSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Face Recognition System")
        self.root.geometry("1200x700")
        self.root.resizable(False, False)
        self.root.configure(bg="#f4f6f9")

        self.base_dir = os.path.dirname(os.path.abspath(__file__))

        self.setup_ui()

    # ---------------- UI SETUP ---------------- #

    def setup_ui(self):

        # Top Banner Frame
        header_frame = Frame(self.root, bg="#1e3d59", height=120)
        header_frame.pack(fill="x")

        title_label = Label(
            header_frame,
            text="FACE RECOGNITION ATTENDANCE SYSTEM",
            font=("Segoe UI", 24, "bold"),
            bg="#1e3d59",
            fg="white"
        )
        title_label.pack(pady=35)

        # Sub Heading
        subtitle = Label(
            self.root,
            text="Smart Attendance Management Using AI",
            font=("Segoe UI", 14),
            bg="#f4f6f9",
            fg="#333"
        )
        subtitle.pack(pady=15)

        # Buttons Frame (Centered)
        button_frame = Frame(self.root, bg="#f4f6f9")
        button_frame.pack(pady=40)

        buttons = [
            ("STUDENT DETAILS", "student.py"),
            ("FACE DETECTOR", "recognizer.py"),
            ("ATTENDANCE", "attendance.py"),
            ("ATTENDANCE SHEET", "sheet.py"),
            ("TRAIN DATA", "train.py"),
            ("EXIT", None),
        ]

        rows = 2
        cols = 3

        for index, (text, script) in enumerate(buttons):
            row = index // cols
            col = index % cols

            if script:
                command = lambda s=script: self.run_script(s)
            else:
                command = self.exit_app

            btn = Button(
                button_frame,
                text=text,
                font=("Segoe UI", 13, "bold"),
                bg="white",
                fg="#1e3d59",
                width=22,
                height=3,
                relief="flat",
                cursor="hand2",
                command=command
            )

            btn.grid(row=row, column=col, padx=30, pady=25)

            # Hover Effect
            btn.bind("<Enter>", lambda e, b=btn: b.configure(bg="#1e3d59", fg="white"))
            btn.bind("<Leave>", lambda e, b=btn: b.configure(bg="white", fg="#1e3d59"))

        # Footer
        footer = Label(
            self.root,
            text="© 2026 Face Recognition System | Developed with Python & OpenCV",
            font=("Segoe UI", 9),
            bg="#f4f6f9",
            fg="gray"
        )
        footer.pack(side="bottom", pady=15)

    # ---------------- SCRIPT RUNNER ---------------- #

    def run_script(self, script_name):
        script_path = os.path.join(self.base_dir, script_name)

        if not os.path.exists(script_path):
            messagebox.showerror("Error", f"{script_name} not found!")
            return

        try:
            subprocess.Popen([sys.executable, script_path])
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------------- EXIT ---------------- #

    def exit_app(self):
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.root.destroy()


# ---------------- MAIN ---------------- #

if __name__ == "__main__":
    root = Tk()
    app = FaceRecognitionSystem(root)
    root.mainloop()
