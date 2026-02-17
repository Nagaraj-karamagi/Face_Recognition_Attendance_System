import cv2
import openpyxl
import os
import numpy as np
from PIL import Image


class FaceRecognitionTrainer:
    def __init__(self, model_path: str, excel_path: str, photo_folder: str):
        self.model_path = model_path
        self.excel_file = excel_path
        self.photo_folder = photo_folder

        self.clf = cv2.face.LBPHFaceRecognizer_create(
            radius=1,
            neighbors=8,
            grid_x=8,
            grid_y=8
        )

        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )

        self.student_data = self.load_student_data()

    # ---------------- LOAD STUDENTS ---------------- #

    def load_student_data(self):
        if not os.path.exists(self.excel_file):
            raise FileNotFoundError("Student Excel file not found")

        wb = openpyxl.load_workbook(self.excel_file)
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

    # ---------------- LOAD DATASET ---------------- #

    def load_dataset(self):
        faces = []
        ids = []

        for student_id in self.student_data.keys():
            student_folder = os.path.join(self.photo_folder, str(student_id))

            if not os.path.isdir(student_folder):
                print(f"Folder missing: {student_folder}")
                continue

            for image_name in os.listdir(student_folder):
                image_path = os.path.join(student_folder, image_name)

                try:
                    pil_image = Image.open(image_path).convert("L")
                except:
                    continue

                img_numpy = np.array(pil_image, dtype="uint8")

                detected_faces = self.face_cascade.detectMultiScale(
                    img_numpy,
                    scaleFactor=1.2,
                    minNeighbors=5,
                    minSize=(50, 50)
                )

                for (x, y, w, h) in detected_faces:
                    face = img_numpy[y:y+h, x:x+w]
                    face = cv2.resize(face, (200, 200))  # Normalize size
                    faces.append(face)
                    ids.append(student_id)

        return faces, ids

    # ---------------- TRAIN ---------------- #

    def train_model(self):
        print("Loading dataset...")
        faces, ids = self.load_dataset()

        if not faces:
            print("No faces found for training.")
            return

        print(f"Training on {len(faces)} face samples...")
        self.clf.train(faces, np.array(ids))
        self.clf.save(self.model_path)

        print(f"Model saved as {self.model_path}")
        print("Training completed successfully ✅")

    # ---------------- EVALUATE ---------------- #

    def evaluate_model(self):
        print("Evaluating model...")

        faces, ids = self.load_dataset()

        if not faces:
            print("No data for evaluation.")
            return

        correct = 0
        total = 0

        for i in range(len(faces)):
            predicted_id, confidence = self.clf.predict(faces[i])

            if predicted_id == ids[i]:
                correct += 1

            total += 1

        accuracy = (correct / total) * 100
        print(f"Model Accuracy: {accuracy:.2f}%")



# ---------------- MAIN ---------------- #

if __name__ == "__main__":
    model_path = "trainer.yml"
    excel_path = "student_data.xlsx"
    photo_folder = "photos"

    try:
        trainer = FaceRecognitionTrainer(
            model_path,
            excel_path,
            photo_folder
        )

        trainer.train_model()
        trainer.evaluate_model()

    except Exception as e:
        print("Error:", e)
