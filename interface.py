import sys
import os
import shutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QDesktopWidget, QLabel
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QIcon
from PyQt5.QtCore import QTimer, QThread, pyqtSignal, Qt
from spreadsheet_unlocker import ExcelWorkbookModifier

class Worker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal()

    def __init__(self, path, parent=None):
        super().__init__(parent)
        self.path = path

    def run(self):
        try:
            if not (self.path.endswith('.xlsx') or self.path.endswith('.xlsm')):
                self.error.emit()
                return

            modifier = ExcelWorkbookModifier(self.path)
            modifier.modify_workbook()
            self.finished.emit(modifier.new_xlsm_file_path)
        except Exception as e:
            self.error.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("SheetUnlock")
        self.setGeometry(100, 100, 600, 400)
        self.setFixedSize(600, 400)
        self.center()
        self.setStyleSheet("background-color: #000749; color: #8cfffb;")
        self.setAcceptDrops(True)

        # Ensure the icon file is in the correct path
        icon_path = 'icon.png'
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"Icon file not found: {icon_path}")

        self.label = QLabel(self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("background-color: #000749; font-family: Times; font-size: 14px; color: #8cfffb;")
        self.label.setGeometry(0, 0, self.width(), self.height())

        self.caret_visible = False
        self.caret_timer = QTimer(self)
        self.caret_timer.timeout.connect(self.blink_caret)
        self.caret_timer.start(500)

        self.append_text("Please drag in your .xlsm or .xlsx file.")

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                self.execute(file_path)
                break

    def execute(self, path):
        self.append_text("Processing file...")
        self.worker = Worker(path)
        self.worker.finished.connect(self.on_file_saved)
        self.worker.error.connect(self.handle_error)
        self.worker.start()

    def handle_error(self):
        self.append_text("Error.")
        QTimer.singleShot(1000, self.reset_greeting)

    def append_text(self, text):
        current_text = self.label.text()
        if current_text and not current_text.endswith('\n'):
            current_text += '\n'
        self.label.setText(current_text + text)

    def blink_caret(self):
        self.caret_visible = not self.caret_visible
        self.update_caret()

    def update_caret(self):
        current_text = self.label.text()
        if self.caret_visible:
            self.label.setText(current_text + '|')
        else:
            self.label.setText(current_text.replace('|', ''))

    def on_file_saved(self, new_file_path):
        self.append_text(f"File saved to: {new_file_path}")
        QTimer.singleShot(2000, self.reset_greeting)

    def reset_greeting(self):
        self.label.clear()
        self.append_text("Please drag in your .xlsm or .xlsx file.")

def main():
    app = QApplication(sys.argv)

    # Set the application name and display name to "SheetUnlock"
    app.setApplicationName("SheetUnlock")
    app.setApplicationDisplayName("SheetUnlock")

    # Ensure the icon file is in the correct path
    icon_path = 'icon.png'
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))  # Set the window icon for the application
        # Set the dock icon for macOS
        if sys.platform == 'darwin':
            from AppKit import NSApplication, NSImage
            NSApplication.sharedApplication().setApplicationIconImage_(NSImage.alloc().initByReferencingFile_(os.path.abspath(icon_path)))
    else:
        print(f"Icon file not found: {icon_path}")

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
