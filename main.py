import sys

from PySide6.QtWidgets import QApplication
from views.main_window import RegistrarApp


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set style
    app.setStyle("Fusion")
    
    window = RegistrarApp()
    window.show()
    
    sys.exit(app.exec())