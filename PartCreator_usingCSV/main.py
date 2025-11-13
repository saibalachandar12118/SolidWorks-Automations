#Author : SAI BALACHANDAR

import sys, os
from pathlib import Path
from PyQt6.QtWidgets import QApplication, QDialog, QLabel, QComboBox, QGroupBox, QFormLayout, QProgressBar, QPushButton, QHBoxLayout, QVBoxLayout
from PyQt6.QtGui import QPixmap , QIcon
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import win32com.client
import csv
import traceback

os.chdir(os.path.dirname(os.path.abspath(__file__)))

class PlateGeneratorThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    message = pyqtSignal(str)
    stopped = False

    def __init__(self, csv_path):
        super().__init__()
        self.csv_path = csv_path

    def run(self):
        def meters(mm): return mm / 1000.0

        csv_path = Path(self.csv_path)
        TEMPLATE_PATH = Path(r"C:\ProgramData\SolidWorks\SOLIDWORKS 2025\templates\Part.prtdot")
        OUTPUT_ROOT = Path(__file__).parent.resolve()

        try:
            sw = win32com.client.Dispatch("SldWorks.Application")
        except Exception:
            sw = win32com.client.DispatchEx("SldWorks.Application")
        sw.Visible = False

        with open(csv_path, "r", newline="") as f:
            reader = list(csv.DictReader(f))
            total = len(reader)
            for i, row in enumerate(reader, 1):
                if self.stopped:
                    self.message.emit("[INFO] Operation cancelled.")
                    break
                try:
                    name = row["Name"].strip()
                    L, W, T = float(row["Length"]), float(row["Width"]), float(row["Thickness"])
                    model = sw.NewDocument(str(TEMPLATE_PATH), 0, 0, 0)
                    model = sw.ActiveDoc
                    sm, fm = model.SketchManager, model.FeatureManager
                    front_plane = model.FeatureByPositionReverse(0)
                    front_plane.Select2(False, 0)
                    sm.InsertSketch(True)
                    sm.CreateCenterRectangle(0, 0, 0, meters(L/2), meters(W/2), 0)
                    sm.InsertSketch(True)
                    fm.FeatureExtrusion2(True, False, False, 0, 0, meters(T), 0,
                                         False, False, False, False, 0, 0,
                                         False, False, False, False, True, True, True, 0, 0, False)
                    sld_path = OUTPUT_ROOT / "Exports" / f"{name}.SLDPRT"
                    step_path = OUTPUT_ROOT / "Exports" / f"{name}.STEP"
                    step_path.parent.mkdir(parents=True, exist_ok=True)
                    model.SaveAs(str(sld_path))
                    model.SaveAs(str(step_path))
                    sw.CloseDoc(model.GetTitle)
                    self.message.emit(f"[DONE] {name}")
                    self.progress.emit(int((i / total) * 100))
                except Exception as e:
                    self.message.emit(f"[ERROR] {e}")
                    traceback.print_exc()
        self.finished.emit()

class PartSelectionDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Part Selection UI")
        self.setMinimumSize(800, 500)
        self.setWindowIcon(QIcon("Window_titleImage.jpg"))        
        self.setStyleSheet("""
            QWidget {
                background-color: #1e1f32;
                color: #e0e0e0;
                font-family: 'Segoe UI';
                font-size: 10pt;
            }
            QGroupBox {
                border: 1px solid #3a3a3a;
                border-radius: 5px;
                margin-top: 6px;
                padding: 6px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                color: #00a0a3;
                font-weight: bold;
            }
            QPushButton {
                background-color: #2b2b2b;
                border: 1px solid #444;
                padding: 6px 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #3b3b3b;
            }
            QPushButton#Apply {
                background-color: #007b7f;
                color: white;
            }
            QProgressBar {
                border: 1px solid #444;
                border-radius: 4px;
                text-align: center;
                height: 18px;
            }
            QProgressBar::chunk {
                background-color: #00a0a3;
            }
        """)

        main_layout = QVBoxLayout()
        top_layout = QHBoxLayout()
        middle_layout = QHBoxLayout()
        bottom_layout = QHBoxLayout()

        self.comboBox = QComboBox()
        csv_dir = Path(r"./../Inputs")
        self.directions = [(str(file.stem), str(file)) for file in csv_dir.rglob("*.csv")]
        for display, value in self.directions:
            self.comboBox.addItem(display, value)
        self.comboBox.setCurrentIndex(0)
        self.label_combobox = QLabel("Select the Model : ")
        top_layout.addWidget(self.label_combobox)
        top_layout.addWidget(self.comboBox)

        Detail_group = QGroupBox("Part Details")
        Chid_formlayout = QFormLayout()
        self.label_name = QLabel(f"{self.comboBox.currentText()}")
        self.label_material = QLabel("Aluminum")
        self.label_mass = QLabel("2.45 kg")
        self.label_dims = QLabel("200 × 150 × 50 mm")
        Chid_formlayout.addRow("Name:", self.label_name)
        Chid_formlayout.addRow("Material:", self.label_material)
        Chid_formlayout.addRow("Mass:", self.label_mass)
        Chid_formlayout.addRow("Dimensions:", self.label_dims)
        Detail_group.setLayout(Chid_formlayout)

        image_group = QGroupBox("Part Image")
        image_layout = QVBoxLayout()
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setStyleSheet("border: 1px solid white; padding: 10px;")
        image_layout.addWidget(self.image_label)
        image_group.setLayout(image_layout)

        middle_layout.addWidget(Detail_group)
        middle_layout.addWidget(image_group)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_apply = QPushButton("Generate Part")
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.setTextVisible(False)
        self.status_label = QLabel("")
        self.btn_cancel.setStyleSheet("background-color: #d32f2f; color: white; border: none; padding: 6px;")
        self.btn_apply.setStyleSheet("background-color: #388e3c; color: white; border: none; padding: 6px;")
        self.progress_bar.setStyleSheet("QProgressBar {background-color: #002b5c; border-radius: 3px;} QProgressBar::chunk {background-color: #00bcd4; width: 20px;}")

        bottom_layout.addWidget(self.progress_bar)
        bottom_layout.addWidget(self.status_label)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_cancel)
        bottom_layout.addWidget(self.btn_apply)

        main_layout.addLayout(top_layout)
        main_layout.addLayout(middle_layout)
        main_layout.addLayout(bottom_layout)
        self.setLayout(main_layout)

        self.update_image_preview()
        self.worker = None
        self.comboBox.currentTextChanged.connect(self.update_image_preview)
        self.btn_apply.clicked.connect(self.start_generation)
        self.btn_cancel.clicked.connect(self.cancel_generation)

    def update_image_preview(self):
        self.file_name = self.comboBox.currentText().removesuffix(".csv") + ".png"
        self.file_name = Path(f'./../Inputs/{self.file_name}').resolve()
        if self.file_name.exists():
            pixmap = QPixmap(str(self.file_name))
            self.image_label.setPixmap(pixmap.scaled(300, 300, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            self.image_label.setText("No Preview Image Found")

    def start_generation(self):
        csv_path = self.comboBox.currentData()
        self.worker = PlateGeneratorThread(csv_path)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.message.connect(self.update_status)
        self.worker.finished.connect(self.on_finished)
        self.progress_bar.setValue(0)
        self.status_label.setText("Processing...")
        self.worker.start()

    def cancel_generation(self):
        if self.worker and self.worker.isRunning():
            self.worker.stopped = True
            self.worker.wait()
            self.status_label.setText("Cancelled by user.")
            self.progress_bar.setValue(0)
            self.close()
        else:
            self.close()


    def update_status(self, msg):
        self.status_label.setText(msg)
        print(msg)

    def on_finished(self):
        self.status_label.setText("Process finished!")
        self.progress_bar.setValue(100)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    dlg = PartSelectionDialog()
    dlg.show()
    dlg.exec()
