import sys
import os
from datetime import datetime, timedelta
from pathlib import Path

from PySide6.QtWidgets import (QWidget, QApplication, QMainWindow, QSplitter, QTreeView, 
                              QTableView, QStatusBar, QMenuBar, QMenu, QFileDialog,
                              QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                              QPushButton, QDateEdit, QTextEdit, QListWidget, 
                              QMessageBox, QAbstractItemView, QInputDialog, QComboBox)
from PySide6.QtGui import QStandardItemModel, QStandardItem, QAction
from PySide6.QtCore import Qt, QDate, QModelIndex
import pandas as pd
from models.document import Document


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


class RegistrarApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Регистратор документов")
        self.resize(1000, 700)
        
        # Initialize data structures
        self.folders = {
            "Входящие": [],
            "Исходящие": [],
            "Договоры": []
        }
        
        # Create UI
        self.create_menus()
        self.create_main_widgets()
        self.create_status_bar()
        
        # Add some sample data
        self.add_sample_data()
    
    def create_menus(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("Файл")
        
        export_action = QAction("Экспорт в Excel", self)
        export_action.triggered.connect(self.export_to_excel)
        file_menu.addAction(export_action)
        
        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Actions menu
        actions_menu = menubar.addMenu("Действия")
        
        check_expiring_action = QAction("Проверить истекающие договора", self)
        check_expiring_action.triggered.connect(self.check_expiring_contracts)
        actions_menu.addAction(check_expiring_action)
        
        search_action = QAction("Поиск", self)
        search_action.triggered.connect(self.show_search_dialog)
        actions_menu.addAction(search_action)
        
        # Help menu
        help_menu = menubar.addMenu("Помощь")
        
        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
        license_action = QAction("Лицензия", self)
        license_action.triggered.connect(self.show_license)
        help_menu.addAction(license_action)
    
    def create_main_widgets(self):
        main_splitter = QSplitter()
        
        # Left side - folder tree
        self.folder_tree = QTreeView()
        self.folder_tree.setHeaderHidden(True)
        self.folder_model = QStandardItemModel()
        self.folder_tree.setModel(self.folder_model)
        
        # Create folder structure
        incoming_item = QStandardItem("Входящие")
        outgoing_item = QStandardItem("Исходящие")
        contracts_item = QStandardItem("Договоры")
        
        self.folder_model.appendRow(incoming_item)
        self.folder_model.appendRow(outgoing_item)
        self.folder_model.appendRow(contracts_item)
        
        # Add some sample subfolders
        incoming_item.appendRow(QStandardItem("2023"))
        incoming_item.appendRow(QStandardItem("2024"))
        
        self.folder_tree.expandAll()
        self.folder_tree.selectionModel().selectionChanged.connect(self.folder_selection_changed)
        
        # Right side - document table and buttons
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        
        # Document table
        self.document_table = QTableView()
        self.document_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.document_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.document_table.doubleClicked.connect(self.view_document)
        
        self.document_model = QStandardItemModel()
        self.document_model.setHorizontalHeaderLabels([
            "Номер", "Наименование", "Контрагент", "Дата начала", "Дата окончания"
        ])
        self.document_table.setModel(self.document_model)
        self.document_table.horizontalHeader().setStretchLastSection(True)
        
        right_layout.addWidget(self.document_table)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        self.add_button = QPushButton("Добавить")
        self.add_button.clicked.connect(self.add_document)
        buttons_layout.addWidget(self.add_button)
        
        self.edit_button = QPushButton("Редактировать")
        self.edit_button.clicked.connect(self.edit_document)
        buttons_layout.addWidget(self.edit_button)
        
        self.delete_button = QPushButton("Удалить")
        self.delete_button.clicked.connect(self.delete_document)
        buttons_layout.addWidget(self.delete_button)
        
        right_layout.addLayout(buttons_layout)
        
        right_widget.setLayout(right_layout)
        
        # Add widgets to splitter
        main_splitter.addWidget(self.folder_tree)
        main_splitter.addWidget(right_widget)
        main_splitter.setSizes([200, 800])
        
        self.setCentralWidget(main_splitter)
    
    def create_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готово")
    
    def add_sample_data(self):
        # Add some sample documents
        today = QDate.currentDate()
        
        doc1 = Document(
            number="INC-2023-001",
            name="Запрос информации",
            counterparty="ООО Контрагент",
            start_date=today.addDays(-30),
            end_date=today.addDays(30),
            description="Запрос информации о поставках",
            attachments=[]
        )
        
        doc2 = Document(
            number="OUT-2023-001",
            name="Ответ на запрос",
            counterparty="ООО Контрагент",
            start_date=today.addDays(-20),
            description="Ответ на запрос информации",
            attachments=[]
        )
        
        doc3 = Document(
            number="CNT-2023-001",
            name="Договор поставки",
            counterparty="ООО Контрагент",
            start_date=today.addDays(-10),
            end_date=today.addDays(20),
            description="Договор на поставку оборудования",
            attachments=[]
        )
        
        self.folders["Входящие"].append(doc1)
        self.folders["Исходящие"].append(doc2)
        self.folders["Договоры"].append(doc3)
        
        # Update table if current folder is one of these
        current_index = self.folder_tree.currentIndex()
        if current_index.isValid():
            self.update_document_table(current_index)
    
    def folder_selection_changed(self, selected, deselected):
        if selected.indexes():
            current_index = selected.indexes()[0]
            self.update_document_table(current_index)
    
    def update_document_table(self, index):
        folder_name = index.data()
        parent = index.parent()
        
        if parent.isValid():
            # This is a subfolder, we'll treat it as empty for this example
            folder_name = parent.data()
            subfolder_name = index.data()
            documents = []
        else:
            # This is a main folder
            documents = self.folders.get(folder_name, [])
        
        self.document_model.clear()
        self.document_model.setHorizontalHeaderLabels([
            "Номер", "Наименование", "Контрагент", "Дата начала", "Дата окончания"
        ])
        
        for doc in documents:
            row = [
                QStandardItem(doc.number),
                QStandardItem(doc.name),
                QStandardItem(doc.counterparty),
                QStandardItem(doc.start_date.toString("dd.MM.yyyy")),
                QStandardItem(doc.end_date.toString("dd.MM.yyyy") if doc.end_date else "")
            ]
            self.document_model.appendRow(row)
        
        self.status_bar.showMessage(f"Папка: {folder_name}. Документов: {len(documents)}")
    
    def get_current_folder_name(self):
        current_index = self.folder_tree.currentIndex()
        if current_index.isValid():
            parent = current_index.parent()
            if parent.isValid():
                return parent.data()
            return current_index.data()
        return ""
    
    def get_selected_document(self):
        selected = self.document_table.selectionModel().selectedRows()
        if not selected:
            return None
        
        row = selected[0].row()
        folder_name = self.get_current_folder_name()
        if folder_name in self.folders:
            if row < len(self.folders[folder_name]):
                return self.folders[folder_name][row]
        return None
    
    def view_document(self, index):
        document = self.get_selected_document()
        if document:
            dialog = DocumentViewDialog(document, self)
            dialog.exec()
    
    def add_document(self):
        folder_name = self.get_current_folder_name()
        if folder_name not in self.folders:
            QMessageBox.warning(self, "Ошибка", "Выберите папку для добавления документа")
            return
        
        dialog = DocumentEditDialog(None, self)
        if dialog.exec() == QDialog.Accepted:
            new_doc = dialog.get_document()
            self.folders[folder_name].append(new_doc)
            self.update_document_table(self.folder_tree.currentIndex())
            self.status_bar.showMessage(f"Документ добавлен: {new_doc.number}")
    
    def edit_document(self):
        document = self.get_selected_document()
        if not document:
            QMessageBox.warning(self, "Ошибка", "Выберите документ для редактирования")
            return
        
        folder_name = self.get_current_folder_name()
        if folder_name not in self.folders:
            return
        
        dialog = DocumentEditDialog(document, self)
        if dialog.exec() == QDialog.Accepted:
            updated_doc = dialog.get_document()
            # Find and replace the document in the folder
            for i, doc in enumerate(self.folders[folder_name]):
                if doc == document:
                    self.folders[folder_name][i] = updated_doc
                    break
            self.update_document_table(self.folder_tree.currentIndex())
            self.status_bar.showMessage(f"Документ обновлен: {updated_doc.number}")
    
    def delete_document(self):
        document = self.get_selected_document()
        if not document:
            QMessageBox.warning(self, "Ошибка", "Выберите документ для удаления")
            return
        
        folder_name = self.get_current_folder_name()
        if folder_name not in self.folders:
            return
        
        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить документ {document.number}?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.folders[folder_name].remove(document)
            self.update_document_table(self.folder_tree.currentIndex())
            self.status_bar.showMessage(f"Документ удален: {document.number}")
    
    def export_to_excel(self):
        folder_name = self.get_current_folder_name()
        if folder_name not in self.folders or not self.folders[folder_name]:
            QMessageBox.warning(self, "Ошибка", "Нет документов для экспорта")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Экспорт в Excel", 
            f"{folder_name}_{QDate.currentDate().toString('yyyyMMdd')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        
        if file_path:
            data = [doc.to_dict() for doc in self.folders[folder_name]]
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)
            self.status_bar.showMessage(f"Документы экспортированы в {file_path}")
    
    def check_expiring_contracts(self):
        today = QDate.currentDate()
        threshold = today.addDays(30)
        expiring = []
        
        for doc in self.folders["Договоры"]:
            if doc.end_date and doc.end_date <= threshold:
                expiring.append(doc)
        
        if expiring:
            message = "Истекающие договоры:\n\n"
            for doc in expiring:
                message += f"{doc.number} - {doc.name} (до {doc.end_date.toString('dd.MM.yyyy')})\n"
            QMessageBox.information(self, "Истекающие договоры", message)
        else:
            QMessageBox.information(self, "Истекающие договоры", "Нет договоров, истекающих в ближайшие 30 дней")
    
    def show_search_dialog(self):
        dialog = SearchDialog(self)
        if dialog.exec() == QDialog.Accepted:
            search_params = dialog.get_search_params()
            self.perform_search(search_params)
    
    def perform_search(self, params):
        results = []
        
        for folder_name, documents in self.folders.items():
            for doc in documents:
                match = False
                
                if params["field"] == "all":
                    # Search in all text fields
                    search_text = params["text"].lower()
                    if (search_text in doc.number.lower() or
                        search_text in doc.name.lower() or
                        search_text in doc.counterparty.lower() or
                        search_text in doc.description.lower()):
                        match = True
                elif params["field"] in ["start_date", "end_date"]:
                    # Date range search
                    date = doc.start_date if params["field"] == "start_date" else doc.end_date
                    if date:
                        from_date = params["from_date"]
                        to_date = params["to_date"]
                        if from_date <= date <= to_date:
                            match = True
                else:
                    # Field-specific text search
                    field_value = getattr(doc, params["field"]).lower()
                    if params["text"].lower() in field_value:
                        match = True
                
                if match:
                    results.append((folder_name, doc))
        
        # Display results in table
        self.document_model.clear()
        self.document_model.setHorizontalHeaderLabels([
            "Папка", "Номер", "Наименование", "Контрагент", "Дата начала", "Дата окончания"
        ])
        
        for folder_name, doc in results:
            row = [
                QStandardItem(folder_name),
                QStandardItem(doc.number),
                QStandardItem(doc.name),
                QStandardItem(doc.counterparty),
                QStandardItem(doc.start_date.toString("dd.MM.yyyy")),
                QStandardItem(doc.end_date.toString("dd.MM.yyyy") if doc.end_date else "")
            ]
            self.document_model.appendRow(row)
        
        self.status_bar.showMessage(f"Найдено документов: {len(results)}")
    
    def show_about(self):
        QMessageBox.about(self, "О программе", 
                         "Регистратор документов\nВерсия 1.0\n\nПрограмма для учета документов в организации")
    
    def show_license(self):
        QMessageBox.information(self, "Лицензия", 
                              "Это программа распространяется под лицензией MIT.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set style
    app.setStyle("Fusion")
    
    window = RegistrarApp()
    window.show()
    
    sys.exit(app.exec())