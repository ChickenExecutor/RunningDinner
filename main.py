import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QInputDialog, QMessageBox
import psutil, time
import os, subprocess, platform

import Helper
import openPairings

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("RunningDinner Calc")
        button = QPushButton("Calculate Pairings")
        get_file = QPushButton('Open Project Folder')
        openPairings = QPushButton('Show Pairings')
        button.setCheckable(True)
        get_file.setCheckable(True)
        openPairings.setCheckable(True)
        button.clicked.connect(self.the_button_was_clicked)
        get_file.clicked.connect(self.getpath)
        openPairings.clicked.connect(self.getpairings)
        layout = QVBoxLayout()
        layout.addWidget(get_file)
        layout.addWidget(button)
        layout.addWidget(openPairings)
        
        # Create a central widget and set the layout
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def getpath(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Select File")

    def getpairings(self):
        team, ok = QInputDialog.getText(self, 'Input Dialog', 'Enter team name:')
        if ok:
            pairings = openPairings.open_pairings(team, self.path)
            msg = f"Possible pairings for team '{team}':\n" + "\n".join(f"- {pairing}" for pairing in pairings)
            QMessageBox.information(self, "Pairings", msg)

    def the_button_was_clicked(self):
        # best-effort: close any open Excel windows before working with files
        try:
            import win32com.client
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                excel.DisplayAlerts = False
                excel.Quit()
            except Exception:
                # no running COM Excel instance or couldn't quit it
                pass
        except Exception:
            # win32com not available
            pass

        # fallback: terminate excel.exe processes if psutil is available
        try:
            for proc in psutil.process_iter(['name']):
                name = (proc.info.get('name') or '').lower()
                if name == 'excel.exe':
                    try:
                        proc.terminate()
                    except Exception:
                        try:
                            proc.kill()
                        except Exception:
                            pass
            time.sleep(0.5)
        except Exception:
            pass
        self.helper = Helper.RunningDinnerHelper(self.path)
        self.helper.clear_paste_area(self.helper.max_team_length)
        self.helper.paste_pairing(self.helper.teams)
        self.helper.sheet.parent.save(self.helper.path)
        print("Pairings calculated and saved.")
        try:
            if hasattr(os, "startfile"):
                os.startfile(self.helper.path)
            else:
                if platform.system() == "Darwin":
                    subprocess.Popen(["open", self.helper.path])
                else:
                    subprocess.Popen(["xdg-open", self.helper.path])
        except Exception:
            pass

app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()