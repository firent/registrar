from PySide6.QtWidgets import (QWidget, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                              QPushButton, QDateEdit, QComboBox)

class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Поиск документов")
        self.resize(400, 200)
        
        layout = QVBoxLayout()
        
        # Search field selection
        self.field_combo = QComboBox()
        self.field_combo.addItem("Все поля", "all")
        self.field_combo.addItem("Номер документа", "number")
        self.field_combo.addItem("Наименование", "name")
        self.field_combo.addItem("Контрагент", "counterparty")
        self.field_combo.addItem("Дата начала", "start_date")
        self.field_combo.addItem("Дата окончания", "end_date")
        self.field_combo.addItem("Описание", "description")
        layout.addWidget(self.field_combo)
        
        # Search text
        self.search_edit = QLineEdit()
        layout.addWidget(QLabel("Текст для поиска:"))
        layout.addWidget(self.search_edit)
        
        # Date range for date fields
        self.date_range_layout = QHBoxLayout()
        self.date_range_layout.addWidget(QLabel("От:"))
        self.from_date_edit = QDateEdit()
        self.from_date_edit.setDisplayFormat("dd.MM.yyyy")
        self.from_date_edit.setCalendarPopup(True)
        self.date_range_layout.addWidget(self.from_date_edit)
        
        self.date_range_layout.addWidget(QLabel("До:"))
        self.to_date_edit = QDateEdit()
        self.to_date_edit.setDisplayFormat("dd.MM.yyyy")
        self.to_date_edit.setCalendarPopup(True)
        self.date_range_layout.addWidget(self.to_date_edit)
        
        self.date_range_widget = QWidget()
        self.date_range_widget.setLayout(self.date_range_layout)
        self.date_range_widget.hide()
        layout.addWidget(self.date_range_widget)
        
        # Connect field change to show/hide date range
        self.field_combo.currentIndexChanged.connect(self.update_field_visibility)
        
        # Search button
        self.search_button = QPushButton("Поиск")
        self.search_button.clicked.connect(self.accept)
        layout.addWidget(self.search_button)
        
        self.setLayout(layout)
    
    def update_field_visibility(self):
        field = self.field_combo.currentData()
        if field in ["start_date", "end_date"]:
            self.date_range_widget.show()
        else:
            self.date_range_widget.hide()
    
    def get_search_params(self):
        field = self.field_combo.currentData()
        text = self.search_edit.text()
        from_date = self.from_date_edit.date() if field in ["start_date", "end_date"] else None
        to_date = self.to_date_edit.date() if field in ["start_date", "end_date"] else None
        
        return {
            "field": field,
            "text": text,
            "from_date": from_date,
            "to_date": to_date
        }