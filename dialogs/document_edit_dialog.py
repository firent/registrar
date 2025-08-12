from PySide6.QtWidgets import (QDialog, QVBoxLayout, QLabel, QPushButton, QTextEdit, 
                               QListWidget, QLineEdit, QDateEdit, QHBoxLayout, QFileDialog)
from PySide6.QtCore import QDate
from models.document import Document


class DocumentEditDialog(QDialog):
    def __init__(self, document=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редактирование документа" if document else "Добавление документа")
        self.resize(500, 400)
        
        self.document = document or Document()
        
        layout = QVBoxLayout()
        
        # Document fields
        fields_layout = QVBoxLayout()
        
        self.number_edit = QLineEdit(self.document.number)
        self.add_field(fields_layout, "Номер документа:", self.number_edit)
        
        self.name_edit = QLineEdit(self.document.name)
        self.add_field(fields_layout, "Наименование:", self.name_edit)
        
        self.counterparty_edit = QLineEdit(self.document.counterparty)
        self.add_field(fields_layout, "Контрагент:", self.counterparty_edit)
        
        self.start_date_edit = QDateEdit(self.document.start_date)
        self.add_field(fields_layout, "Дата начала:", self.start_date_edit)
        
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDisplayFormat("dd.MM.yyyy")
        self.end_date_edit.setCalendarPopup(True)
        if self.document.end_date:
            self.end_date_edit.setDate(self.document.end_date)
        else:
            self.end_date_edit.setSpecialValueText("Не указана")
            self.end_date_edit.setDate(QDate())
        self.add_field(fields_layout, "Дата окончания:", self.end_date_edit)
        
        layout.addLayout(fields_layout)
        
        # Description
        layout.addWidget(QLabel("Описание:"))
        self.description_edit = QTextEdit(self.document.description)
        layout.addWidget(self.description_edit)
        
        # Attachments
        layout.addWidget(QLabel("Приложенные документы:"))
        self.attachments_list = QListWidget()
        self.attachments_list.addItems(self.document.attachments)
        layout.addWidget(self.attachments_list)
        
        # Buttons for attachments
        buttons_layout = QHBoxLayout()
        self.add_attachment_button = QPushButton("Добавить")
        self.add_attachment_button.clicked.connect(self.add_attachment)
        buttons_layout.addWidget(self.add_attachment_button)
        
        self.remove_attachment_button = QPushButton("Удалить")
        self.remove_attachment_button.clicked.connect(self.remove_attachment)
        buttons_layout.addWidget(self.remove_attachment_button)
        
        layout.addLayout(buttons_layout)
        
        # Dialog buttons
        buttons_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        buttons_layout.addWidget(self.ok_button)
        
        self.cancel_button = QPushButton("Отмена")
        self.cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(self.cancel_button)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
    
    def add_field(self, layout, label, widget):
        field_layout = QHBoxLayout()
        field_layout.addWidget(QLabel(label))
        field_layout.addWidget(widget)
        layout.addLayout(field_layout)
    
    def add_attachment(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл")
        if file_path:
            self.attachments_list.addItem(file_path)
    
    def remove_attachment(self):
        current_item = self.attachments_list.currentItem()
        if current_item:
            self.attachments_list.takeItem(self.attachments_list.row(current_item))
    
    def get_document(self):
        self.document.number = self.number_edit.text()
        self.document.name = self.name_edit.text()
        self.document.counterparty = self.counterparty_edit.text()
        self.document.start_date = self.start_date_edit.date()
        self.document.end_date = self.end_date_edit.date() if not self.end_date_edit.date().isNull() else None
        self.document.description = self.description_edit.toPlainText()
        self.document.attachments = [self.attachments_list.item(i).text() 
                                    for i in range(self.attachments_list.count())]
        return self.document