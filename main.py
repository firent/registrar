import sys
import os
import shutil
import sqlite3
import shortuuid
import pandas as pd

from PySide6.QtWidgets import (QApplication, QWidget, QMainWindow, QSplitter, QTreeView,
                               QTableView, QStatusBar, QFileDialog, QDialog,
                               QVBoxLayout, QHBoxLayout, QInputDialog,
                               QPushButton, QMessageBox, QAbstractItemView,
                               QLabel, QTextEdit, QListWidget, QLineEdit,
                               QDateEdit, QComboBox)
from PySide6.QtGui import QStandardItemModel, QStandardItem, QAction, QColor
from PySide6.QtCore import QDate


class Document:
    def __init__(self, number="", name="", counterparty="",
                 start_date=None, end_date=None,
                 description="", attachments=None, db_id=None):
        self.id = db_id # Идентификатор в БД
        self.number = number
        self.name = name
        self.counterparty = counterparty
        # Убедимся, что даты - это QDate
        if isinstance(start_date, str):
            self.start_date = QDate.fromString(start_date, "dd.MM.yyyy")
        elif isinstance(start_date, QDate) or start_date is None:
            self.start_date = start_date or QDate.currentDate()
        else:
            self.start_date = QDate.currentDate()

        if isinstance(end_date, str):
            self.end_date = QDate.fromString(end_date, "dd.MM.yyyy") if end_date else None
        elif isinstance(end_date, QDate) or end_date is None:
            self.end_date = end_date
        else:
             self.end_date = None

        self.description = description
        self.attachments = attachments or [] # Список путей к файлам в папке attachments

    def is_document_expiring(self, days_threshold=30):
        if not self.end_date:
            return False
        days_left = QDate.currentDate().daysTo(self.end_date)
        # print(f"Document {self.number}: Days left = {days_left}, Threshold = {days_threshold}") # Debug
        return 0 <= days_left <= days_threshold # Истекающий = осталось от 0 до threshold дней

    def to_dict(self):
        return {
            "number": self.number,
            "name": self.name,
            "counterparty": self.counterparty,
            "start_date": self.start_date.toString("dd.MM.yyyy"),
            "end_date": self.end_date.toString("dd.MM.yyyy") if self.end_date else "",
            "description": self.description,
            "attachments": ", ".join(self.attachments) # Сохраняем пути как строку
        }


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
        self.from_date_edit.setDate(QDate.currentDate().addMonths(-1)) # Default range
        self.date_range_layout.addWidget(self.from_date_edit)
        self.date_range_layout.addWidget(QLabel("До:"))
        self.to_date_edit = QDateEdit()
        self.to_date_edit.setDisplayFormat("dd.MM.yyyy")
        self.to_date_edit.setCalendarPopup(True)
        self.to_date_edit.setDate(QDate.currentDate())
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
        # Получаем путь к файлу относительно папки attachments
        file_name = item.text()
        # Предполагаем, что файл хранится в папке attachments рядом с exe
        attachments_dir = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "attachments")
        file_path = os.path.join(attachments_dir, file_name)

        if os.path.exists(file_path):
            try:
                # Используем QDesktopServices для более кроссплатформенного открытия
                from PySide6.QtGui import QDesktopServices
                from PySide6.QtCore import QUrl
                QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть файл: {e}\nПуть: {file_path}")
        else:
            QMessageBox.warning(self, "Ошибка", f"Файл не найден: {file_path}")


class DocumentEditDialog(QDialog):
    def __init__(self, document=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редактирование документа" if document else "Добавление документа")
        self.resize(500, 400)
        self.document = document or Document()
        # Определяем путь к папке attachments
        self.attachments_dir = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "attachments")
        os.makedirs(self.attachments_dir, exist_ok=True) # Создаем папку, если её нет

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
        self.start_date_edit.setDisplayFormat("dd.MM.yyyy")
        self.start_date_edit.setCalendarPopup(True)
        self.add_field(fields_layout, "Дата начала:", self.start_date_edit)

        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDisplayFormat("dd.MM.yyyy")
        self.end_date_edit.setCalendarPopup(True)
        if self.document.end_date:
            self.end_date_edit.setDate(self.document.end_date)
        else:
            self.end_date_edit.setSpecialValueText("Не указана")
            self.end_date_edit.setDate(self.end_date_edit.minimumDate()) # Используем минимальную дату как "пустую"
        self.add_field(fields_layout, "Дата окончания:", self.end_date_edit)
        layout.addLayout(fields_layout)

        # Description
        layout.addWidget(QLabel("Описание:"))
        self.description_edit = QTextEdit(self.document.description)
        layout.addWidget(self.description_edit)

        # Attachments
        layout.addWidget(QLabel("Приложенные документы:"))
        self.attachments_list = QListWidget()
        # Отображаем только имена файлов, не полные пути
        self.attachments_list.addItems([os.path.basename(f) for f in self.document.attachments])
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
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Выберите файл(ы)")
        if file_paths:
            for file_path in file_paths:
                if os.path.exists(file_path):
                    # Генерируем уникальное имя файла
                    original_name = os.path.basename(file_path)
                    name, ext = os.path.splitext(original_name)
                    unique_name = f"{name}_{shortuuid.uuid()}{ext}"

                    # Копируем файл в папку attachments
                    destination_path = os.path.join(self.attachments_dir, unique_name)
                    try:
                        shutil.copy2(file_path, destination_path)
                        # Добавляем только имя файла в список (пути хранятся относительно attachments)
                        self.attachments_list.addItem(unique_name)
                    except Exception as e:
                        QMessageBox.warning(self, "Ошибка", f"Не удалось скопировать файл {original_name}: {e}")

    def remove_attachment(self):
        current_item = self.attachments_list.currentItem()
        if current_item:
            file_name = current_item.text()
            # Удаляем файл из папки attachments (опционально, можно оставить)
            # file_path = os.path.join(self.attachments_dir, file_name)
            # if os.path.exists(file_path):
            #     try:
            #         os.remove(file_path)
            #     except Exception as e:
            #         QMessageBox.warning(self, "Ошибка", f"Не удалось удалить файл {file_name}: {e}")

            # Удаляем элемент из списка
            self.attachments_list.takeItem(self.attachments_list.row(current_item))

    def get_document(self):
        self.document.number = self.number_edit.text()
        self.document.name = self.name_edit.text()
        self.document.counterparty = self.counterparty_edit.text()
        self.document.start_date = self.start_date_edit.date()
        # Проверяем, была ли установлена дата окончания или она "пустая"
        if self.end_date_edit.date() == self.end_date_edit.minimumDate():
             self.document.end_date = None
        else:
             self.document.end_date = self.end_date_edit.date()

        self.document.description = self.description_edit.toPlainText()
        # Получаем список имен файлов из QListWidget
        self.document.attachments = [self.attachments_list.item(i).text()
                                    for i in range(self.attachments_list.count())]
        return self.document


class RegistrarApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Регистратор документов")
        self.resize(1000, 700)

        # Определяем пути к БД и папке вложений
        self.db_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "documents.db")
        self.attachments_dir = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "attachments")
        os.makedirs(self.attachments_dir, exist_ok=True)

        # Инициализируем базу данных
        self.initialize_database()

        # Initialize data structures
        self.folders = {
            "Входящие": [],
            "Исходящие": [],
            "Договоры": []
        }

        # Create status bar
        self.create_status_bar()
            
        # Load data from DB
        self.load_data_from_db()

        # Create UI
        self.create_menus()
        self.create_main_widgets()

        # Вызываем очистку после загрузки данных
        self.cleanup_orphaned_attachments()


    def initialize_database(self):
        """Создает таблицы в базе данных, если они не существуют."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                folder TEXT NOT NULL,
                number TEXT NOT NULL,
                name TEXT NOT NULL,
                counterparty TEXT,
                start_date TEXT NOT NULL, -- Храним как строку dd.MM.yyyy
                end_date TEXT,            -- Храним как строку dd.MM.yyyy или NULL
                description TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                file_path TEXT NOT NULL, -- Храним путь относительно папки attachments
                FOREIGN KEY (document_id) REFERENCES documents (id) ON DELETE CASCADE
            )
        ''')
        conn.commit()
        conn.close()

    def save_data_to_db(self):
        """Сохраняет данные из памяти в базу данных."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Очистка существующих данных (простой способ, можно сделать более умное обновление)
        cursor.execute('DELETE FROM attachments')
        cursor.execute('DELETE FROM documents')

        for folder_name, documents in self.folders.items():
            for doc in documents:
                # Вставляем документ
                cursor.execute('''
                    INSERT INTO documents (folder, number, name, counterparty, start_date, end_date, description)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (folder_name, doc.number, doc.name, doc.counterparty,
                      doc.start_date.toString("dd.MM.yyyy"),
                      doc.end_date.toString("dd.MM.yyyy") if doc.end_date else None,
                      doc.description))
                doc.id = cursor.lastrowid # Получаем ID вставленного документа

                # Вставляем вложения
                for file_path in doc.attachments:
                    cursor.execute('''
                        INSERT INTO attachments (document_id, file_path)
                        VALUES (?, ?)
                    ''', (doc.id, file_path)) # file_path уже относительный путь

        conn.commit()
        conn.close()
        self.status_bar.showMessage("Данные сохранены в базу данных.")

    def load_data_from_db(self):
        """Загружает данные из базы данных в память."""
        self.folders = {key: [] for key in self.folders} # Очищаем текущие данные
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Загружаем документы
        cursor.execute('SELECT * FROM documents')
        rows = cursor.fetchall()
        doc_id_map = {} # Словарь для сопоставления DB ID с объектами Document
        for row in rows:
            doc_id, folder, number, name, counterparty, start_date_str, end_date_str, description = row
            doc = Document(
                number=number,
                name=name,
                counterparty=counterparty,
                start_date=start_date_str, # Конструктор Document обработает строку
                end_date=end_date_str,     # Конструктор Document обработает строку или None
                description=description,
                attachments=[], # Вложения загрузим отдельно
                db_id=doc_id
            )
            if folder in self.folders:
                self.folders[folder].append(doc)
                doc_id_map[doc_id] = doc # Сохраняем ссылку на объект

        # Загружаем вложения
        cursor.execute('SELECT document_id, file_path FROM attachments')
        attachment_rows = cursor.fetchall()
        for doc_id, file_path in attachment_rows:
            if doc_id in doc_id_map:
                doc_id_map[doc_id].attachments.append(file_path) # file_path уже относительный

        conn.close()
        self.status_bar.showMessage("Данные загружены из базы данных.")

    def cleanup_orphaned_attachments(self):
        """Удаляет файлы из папки attachments, которые больше не связаны с документами."""
        # Получаем список всех файлов, связанных с документами
        connected_files = set()
        for docs in self.folders.values():
            for doc in docs:
                connected_files.update(doc.attachments)

        # Получаем список всех файлов в папке attachments
        if os.path.exists(self.attachments_dir):
            all_files = set(os.listdir(self.attachments_dir))
            # Находим файлы, которые не связаны
            orphaned_files = all_files - connected_files

            # Удаляем осиротевшие файлы
            for file_name in orphaned_files:
                file_path = os.path.join(self.attachments_dir, file_name)
                try:
                    os.remove(file_path)
                    print(f"Удален осиротевший файл: {file_path}") # Можно заменить на логирование
                except Exception as e:
                    print(f"Ошибка при удалении файла {file_path}: {e}") # Можно заменить на логирование

            if orphaned_files:
                self.status_bar.showMessage(f"Удалено {len(orphaned_files)} осиротевших файлов.")


    def closeEvent(self, event):
        """Переопределяем событие закрытия окна для сохранения данных."""
        self.save_data_to_db()
        self.cleanup_orphaned_attachments() # Очищаем при выходе
        event.accept()

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
        self.folder_tree.expandAll()
        self.folder_tree.selectionModel().selectionChanged.connect(self.folder_selection_changed)

        # Right side - document table and buttons
        right_widget = QWidget()
        right_layout = QVBoxLayout()

        # Document table
        self.document_table = QTableView()
        self.document_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.document_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
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

    def folder_selection_changed(self, selected, deselected):
        if selected.indexes():
            current_index = selected.indexes()[0]
            self.update_document_table(current_index)

    def update_document_table(self, index, mode="normal"):
        """Обновляет таблицу документов
        :param index: Индекс выбранной папки
        :param mode: Режим отображения ('normal' или 'expiring')
        """
        if mode == "expiring":
            return  # В этом случае используем show_expiring_contracts

        folder_name = index.data()
        parent = index.parent()
        if parent.isValid():
            folder_name = parent.data()
            # subfolder_name = index.data()
            # documents = [] # TODO: Обработка подпапок, если нужно
        documents = self.folders.get(folder_name, [])

        self.document_model.clear()
        headers = ["Номер", "Наименование", "Контрагент", "Дата начала", "Дата окончания"]
        self.document_model.setHorizontalHeaderLabels(headers)
        for doc in documents:
            end_date_str = doc.end_date.toString("dd.MM.yyyy") if doc.end_date else "Бессрочный"
            row = [
                QStandardItem(doc.number),
                QStandardItem(doc.name),
                QStandardItem(doc.counterparty),
                QStandardItem(doc.start_date.toString("dd.MM.yyyy")),
                QStandardItem(end_date_str)
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
        if dialog.exec() == QDialog.DialogCode.Accepted:
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
        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated_doc = dialog.get_document()
            # Найдем и заменим документ в папке
            for i, doc in enumerate(self.folders[folder_name]):
                if doc.id == updated_doc.id: # Используем ID из БД
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
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.folders[folder_name].remove(document)
            # Удаляем связанные файлы из папки attachments
            for file_name in document.attachments:
                file_path = os.path.join(self.attachments_dir, file_name)
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        print(f"Ошибка при удалении файла {file_path}: {e}") # Можно заменить на логирование

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
            minValue=1,
            maxValue=365
        )
        if ok:
            self.show_expiring_contracts(days)

    def show_expiring_contracts(self, days_threshold=30):
        """Показать истекающие договоры в правой панели"""
        expiring_docs = []
        # Собираем все истекающие договоры
        for doc in self.folders["Договоры"]:
            if doc.is_document_expiring(days_threshold):
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
                try:
                    days_left = int(days_item.text())
                except ValueError:
                    continue # Пропускаем, если не число

                color = QColor(0, 0, 0)  # По умолчанию черный
                if days_left <= 7:
                    color = QColor(255, 0, 0)  # Красный для срочных (0-7 дней)
                elif days_left <= 14:
                    color = QColor(255, 165, 0)  # Оранжевый (8-14 дней)

                # Подсвечиваем всю строку
                for col in range(self.document_model.columnCount()):
                    item = self.document_model.item(row, col)
                    if item:
                        item.setForeground(color)

    def show_search_dialog(self):
        dialog = SearchDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            search_params = dialog.get_search_params()
            self.perform_search(search_params)

    def perform_search(self, params):
        # if params.get("special_search") == "expiring_contracts": # Эта часть кода была лишней
        #     self.show_expiring_contracts(params.get("days_threshold", 30))
        #     return
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
                        # print(f"Searching date: {date}, From: {from_date}, To: {to_date}") # Debug
                        if from_date <= date <= to_date:
                            match = True
                else:
                    # Field-specific text search
                    field_value = getattr(doc, params["field"], "")
                    if isinstance(field_value, str) and params["text"].lower() in field_value.lower():
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
                         "Регистратор документов\nВерсия 1.1\nПрограмма для учета документов в организации.\n\nИзменения:\n- Добавлена база данных SQLite для хранения данных.\n- Добавлено управление прикрепленными файлами.")

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
