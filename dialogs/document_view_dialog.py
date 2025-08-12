import os

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QLabel, QPushButton, QTextEdit, 
                               QListWidget, QMessageBox)


class DocumentViewDialog(QDialog):
    def __init__(self, document, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Просмотр документа")
        self.resize(600, 500)
        
        self.document = document
        
        layout = QVBoxLayout()
        
        # Document fields
        self.number_label = QLabel(f"Номер: {document.number}")
        self.name_label = QLabel(f"Наименование: {document.name}")
        self.counterparty_label = QLabel(f"Контрагент: {document.counterparty}")
        self.start_date_label = QLabel(f"Дата начала: {document.start_date.toString('dd.MM.yyyy')}")
        self.end_date_label = QLabel(f"Дата окончания: {document.end_date.toString('dd.MM.yyyy') if document.end_date else 'Не указана'}")
        
        layout.addWidget(self.number_label)
        layout.addWidget(self.name_label)
        layout.addWidget(self.counterparty_label)
        layout.addWidget(self.start_date_label)
        layout.addWidget(self.end_date_label)
        
        # Description
        layout.addWidget(QLabel("Описание:"))
        self.description_text = QTextEdit()
        self.description_text.setPlainText(document.description)
        self.description_text.setReadOnly(True)
        layout.addWidget(self.description_text)
        
        # Attachments
        layout.addWidget(QLabel("Приложенные документы:"))
        self.attachments_list = QListWidget()
        self.attachments_list.addItems(document.attachments)
        self.attachments_list.itemDoubleClicked.connect(self.open_attachment)
        layout.addWidget(self.attachments_list)
        
        # Close button
        self.close_button = QPushButton("Закрыть")
        self.close_button.clicked.connect(self.accept)
        layout.addWidget(self.close_button)
        
        self.setLayout(layout)
    
    def open_attachment(self, item):
        file_path = item.text()
        if os.path.exists(file_path):
            os.startfile(file_path)
        else:
            QMessageBox.warning(self, "Ошибка", f"Файл не найден: {file_path}")