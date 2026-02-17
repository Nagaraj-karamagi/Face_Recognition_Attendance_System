import cv2
import openpyxl
import os
import sys
import tkinter as tk
from tkinter import messagebox


class FaceRecognition:
    def __init__(self, model_path: str, excel_path: str):
        if not os.path.exists(model_path):
            raise FileNotFoundError(f"Model file not found: {model_path}")

        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        self.clf = cv2.face.LBPHFaceRecognizer_create()
        self.clf.read(model_path)

        self.student_data = self.load_student_data(excel_path)

        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )

    # ---------------- LOAD STUDENT DATA ---------------- #

    def load_student_data(self, excel_path):
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        data = {}

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            try:
                student_id = int(row[0])
                name = str(row[1])
                data[student_id] = name
            except:
                continue

        return data

    # ---------------- RECOGNITION ---------------- #

    def recognize_faces(self):
        cap = cv2.VideoCapture(0)

        if not cap.isOpened():
            messagebox.showerror("Error", "Camera not accessible")
            return

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

            faces = self.face_cascade.detectMultiScale(
                gray,
                scaleFactor=1.2,
                minNeighbors=5,
                minSize=(50, 50)
            )

            for (x, y, w, h) in faces:
                face = gray[y:y+h, x:x+w]
                face = cv2.resize(face, (200, 200))

                try:
                    face_id, confidence = self.clf.predict(face)
                except cv2.error:
                    continue

                name = self.student_data.get(face_id, "Unknown")

                # Confidence logic (lower = better in LBPH)
                if confidence < 60:
                    color = (0, 255, 0)
                    label = f"{name} ({round(100 - confidence)}%)"
                else:
                    color = (0, 0, 255)
                    label = "Unknown"

                cv2.rectangle(frame, (x, y), (x+w, y+h), color, 2)
                cv2.putText(
                    frame,
                    label,
                    (x, y - 10),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.8,
                    color,
                    2
                )

            cv2.imshow("Face Recognition", frame)

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()


# ---------------- MAIN ---------------- #

def start_recognition():
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        model_path = os.path.join(base_dir, "trainer.yml")
        excel_path = os.path.join(base_dir, "student_data.xlsx")

        fr = FaceRecognition(model_path, excel_path)
        fr.recognize_faces()

    except Exception as e:
        messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Face Recognition")
    root.geometry("300x150")

    start_button = tk.Button(
        root,
        text="Start Face Recognition",
        font=("Arial", 12),
        command=start_recognition
    )
    start_button.pack(pady=40)

    root.mainloop()
