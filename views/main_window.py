import pandas as pd

from PySide6.QtWidgets import (QWidget, QMainWindow, QSplitter, QTreeView, 
                              QTableView, QStatusBar, QFileDialog,
                              QDialog, QVBoxLayout, QHBoxLayout,
                              QPushButton, QMessageBox, QAbstractItemView)
from PySide6.QtGui import QStandardItemModel, QStandardItem, QAction
from PySide6.QtCore import QDate, Qt
from models.document import Document
from dialogs.document_view_dialog import DocumentViewDialog
from dialogs.document_edit_dialog import DocumentEditDialog
from dialogs.search_dialog import SearchDialog
from utils.date_utils import is_document_expiring, format_days_left


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
        
        check_expiring_action = QAction("Проверить истекающие договоры", self)
        check_expiring_action.triggered.connect(self.show_expiring_contracts_dialog)
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
    
    def show_expiring_contracts_dialog(self):
        """Диалог для выбора порогового значения и показ результатов"""
        days, ok = QInputDialog.getInt(
            self,
            "Проверка истекающих договоров",
            "Введите пороговое количество дней:",
            value=30,
            min=1,
            max=365
        )
    
        if ok:
            self.show_expiring_contracts(days)

    def show_expiring_contracts(self, days_threshold=30):
        """Показать истекающие договоры в правой панели"""
        expiring_docs = []
    
        # Собираем все истекающие договоры
        for doc in self.folders["Договоры"]:
            if doc.end_date and self.is_document_expiring(doc, days_threshold):
                expiring_docs.append(doc)
    
        # Очищаем текущую таблицу
        self.document_model.clear()
        self.document_model.setHorizontalHeaderLabels([
            "Номер", "Наименование", "Контрагент", "Дата окончания", "Дней осталось"
        ])

        # Заполняем таблицу
        for doc in expiring_docs:
            days_left = QDate.currentDate().daysTo(doc.end_date)
            row = [
                QStandardItem(doc.number),
                QStandardItem(doc.name),
                QStandardItem(doc.counterparty),
                QStandardItem(doc.end_date.toString("dd.MM.yyyy")),
                QStandardItem(str(days_left))
            ]
            self.document_model.appendRow(row)
    
        # Обновляем статус бар
        count = len(expiring_docs)
        self.status_bar.showMessage(
            f"Найдено истекающих договоров: {count}. " 
            f"Порог: {days_threshold} дней"
        )
    
        # Подсветка строк с малым сроком
        self.highlight_expiring_contracts()

    def highlight_expiring_contracts(self):
        """Подсветка строк в зависимости от оставшегося срока"""
        for row in range(self.document_model.rowCount()):
            days_item = self.document_model.item(row, 4)  # Столбец "Дней осталось"
            if days_item:
                days_left = int(days_item.text())
                color = Qt.black  # По умолчанию
            
                if days_left <= 7:
                    color = Qt.red
                elif days_left <= 14:
                    color = QColor(255, 165, 0)  # Оранжевый
            
                days_item.setForeground(color)
                # Можно также подсветить всю строку:
                for col in range(self.document_model.columnCount()):
                    item = self.document_model.item(row, col)
                    if item:
                        item.setForeground(color)
    
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
        