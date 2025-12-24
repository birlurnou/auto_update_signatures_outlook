import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import configparser
import pyodbc


class DatabaseManager:
    def __init__(self):
        config = configparser.ConfigParser()
        config.read('config.ini')
        self.SQL_SERVER = config['database']['SQL_SERVER']
        self.SQL_DB = config['database']['SQL_DB']
        self.SQL_USER = config['database']['SQL_USER']
        self.SQL_PASSWORD = config['database']['SQL_PASSWORD']
        self.conn_str = f'DRIVER={{SQL Server}};SERVER={self.SQL_SERVER};DATABASE={self.SQL_DB};UID={self.SQL_USER};PWD={self.SQL_PASSWORD}'

    def get_connection(self):
        try:
            return pyodbc.connect(self.conn_str)
        except pyodbc.Error as e:
            print(f'Connection error: {e}')
            return None

    def execute_query(self, query, params=None, fetch_all=False):
        conn = self.get_connection()
        if not conn:
            return []
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            if fetch_all:
                return cursor.fetchall()
            return []
        except pyodbc.Error as e:
            print(f'Query error: {e}')
            return []
        finally:
            cursor.close()
            conn.close()

    def get_all_users(self):
        query = 'SELECT signature_id, global_id, signature_name, first_name, last_name, email FROM signatures'
        return self.execute_query(query, fetch_all=True)

    def search_users(self, search_term):
        query = '''
        SELECT signature_id, global_id, signature_name, first_name, last_name, email 
        FROM signatures 
        WHERE global_id LIKE ? OR signature_name LIKE ? OR first_name LIKE ? 
        OR last_name LIKE ? OR email LIKE ?
        '''
        pattern = f'%{search_term}%'
        return self.execute_query(query, [pattern] * 5, fetch_all=True)

    def get_user_by_id(self, signature_id):
        query = 'SELECT * FROM signatures WHERE signature_id = ?'
        result = self.execute_query(query, [signature_id], fetch_all=True)
        if result:
            print(f"DEBUG: Found user, row length: {len(result[0])}")
        return result


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.initUI()
        self.load_data()
        # Изначально деактивируем кнопки, требующие выбора
        self.update_button_states()

    def initUI(self):
        self.setWindowTitle('Signature Management')
        self.setGeometry(100, 100, 1000, 600)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # Левая часть
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(10)

        # Поиск
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMinimumHeight(30)
        self.search_input.returnPressed.connect(self.on_search)

        search_btn = QPushButton("Search")
        search_btn.setFixedWidth(80)
        search_btn.clicked.connect(self.on_search)

        search_layout.addWidget(self.search_input)
        search_layout.addWidget(search_btn)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(['ID', 'Global ID', 'Sig name', 'First name', 'Last name', 'Email'])
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.doubleClicked.connect(self.on_double_click)
        # Подключаем сигнал изменения выделения
        self.table.itemSelectionChanged.connect(self.on_selection_changed)

        # Стиль таблицы
        palette = self.table.palette()
        palette.setColor(palette.Highlight, QColor(0, 120, 215))
        palette.setColor(palette.HighlightedText, Qt.white)
        self.table.setPalette(palette)

        left_layout.addLayout(search_layout)
        left_layout.addWidget(self.table)

        # Правая часть с кнопками
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setAlignment(Qt.AlignTop)
        right_layout.setSpacing(10)

        self.create_btn = QPushButton("Create")
        self.copy_btn = QPushButton("Copy")  # Новая кнопка
        self.edit_btn = QPushButton("Edit")
        self.more_settings_btn = QPushButton("More settings")  # Новая кнопка
        self.delete_btn = QPushButton("Delete")

        buttons = [self.create_btn, self.copy_btn, self.edit_btn, self.more_settings_btn, self.delete_btn]
        for btn in buttons:
            btn.setMinimumHeight(35)
            btn.setMinimumWidth(150)
            right_layout.addWidget(btn)

        right_layout.addStretch()

        # Подключаем кнопки
        self.create_btn.clicked.connect(self.on_create)
        self.copy_btn.clicked.connect(self.on_copy)  # Новая функция
        self.edit_btn.clicked.connect(self.on_edit)
        self.more_settings_btn.clicked.connect(self.on_more_settings)  # Новая функция
        self.delete_btn.clicked.connect(self.on_delete)

        layout.addWidget(left_widget, 3)
        layout.addWidget(right_widget, 1)

    def load_data(self, data=None):
        if data is None:
            data = self.db.get_all_users()

        # print(f"DEBUG: Data length: {len(data) if data else 0}")
        # if data:
        #     print(f"DEBUG: First row length: {len(data[0])}")
        #     print(f"DEBUG: First row: {data[0]}")

        self.table.setRowCount(0)
        if data:
            self.table.setRowCount(len(data))
            for row_idx, row in enumerate(data):
                for col_idx, value in enumerate(row):
                    item = QTableWidgetItem(str(value) if value else "")
                    self.table.setItem(row_idx, col_idx, item)

        self.table.resizeColumnsToContents()
        # Обновляем состояние кнопок после загрузки данных
        self.update_button_states()

    def update_button_states(self):
        """Обновляет состояние кнопок в зависимости от выбора в таблице"""
        has_selection = self.table.currentRow() >= 0

        # Кнопки, которые требуют выбора элемента
        self.copy_btn.setEnabled(has_selection)
        self.edit_btn.setEnabled(has_selection)
        self.delete_btn.setEnabled(has_selection)

        # Кнопка Create всегда активна
        self.create_btn.setEnabled(True)
        self.more_settings_btn.setEnabled(True)

    def on_selection_changed(self):
        """Обработчик изменения выделения в таблице"""
        self.update_button_states()

    def on_search(self):
        term = self.search_input.text().strip()
        data = self.db.search_users(term) if term else self.db.get_all_users()
        self.load_data(data)

    def on_create(self):
        dialog = SimpleEditDialog(None, self.db, self)
        if dialog.exec_():
            self.on_search()

    def on_edit(self):
        row = self.table.currentRow()
        if row >= 0:
            signature_id = self.table.item(row, 0).text()
            dialog = SimpleEditDialog(signature_id, self.db, self)
            if dialog.exec_():
                self.on_search()
        else:
            QMessageBox.warning(self, "Warning", "Select a signature first")

    def on_copy(self):
        row = self.table.currentRow()
        if row >= 0:
            signature_id = self.table.item(row, 0).text()
            # Реализуй логику копирования
            QMessageBox.information(self, "Copy", "Copy functionality - to be implemented")
        else:
            QMessageBox.warning(self, "Warning", "Select a signature first")

    def on_more_settings(self):
        QMessageBox.information(self, "More Settings", "More settings functionality - to be implemented")

    def on_delete(self):
        row = self.table.currentRow()
        if row >= 0:
            reply = QMessageBox.question(self, "Confirm", "Delete this signature?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                global_id = self.table.item(row, 0).text()
                print(f"Deleting signature {global_id}")
        else:
            QMessageBox.warning(self, "Warning", "Select a signature first")

    def on_double_click(self, index):
        row = index.row()
        signature_id = self.table.item(row, 0).text()
        dialog = SimpleEditDialog(signature_id, self.db, self)
        if dialog.exec_():
            self.on_search()


class SimpleEditDialog(QDialog):
    def __init__(self, signature_id=None, db=None, parent=None):
        super().__init__(parent)
        self.signature_id = signature_id
        self.db = db
        self.initUI()
        if signature_id:
            self.load_user_data()

    def initUI(self):
        title = "Edit Signature" if self.signature_id else "Create Signature"
        self.setWindowTitle(title)
        self.setMinimumWidth(1000)
        self.setMinimumHeight(600)

        # Основной layout
        main_layout = QVBoxLayout()

        # Scroll area чтобы всё поместилось
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)

        # Контейнер с двумя колонками
        columns_widget = QWidget()
        columns_layout = QHBoxLayout(columns_widget)
        columns_layout.setSpacing(20)

        # Левая колонка
        left_column = QWidget()
        left_layout = QVBoxLayout(left_column)
        left_layout.setSpacing(10)

        # Правая колонка
        right_column = QWidget()
        right_layout = QVBoxLayout(right_column)
        right_layout.setSpacing(10)

        # Создаём все поля заранее
        self.inputs = {}

        # ЛЕВАЯ КОЛОНКА - Основные поля
        left_group = QGroupBox("Basic Information")
        left_group_layout = QFormLayout()
        left_group_layout.setSpacing(8)

        basic_fields = [
            ("global_id", "Global ID"),
            ("signature_name", "Signature Name"),
            ("first_name", "First Name"),
            ("last_name", "Last Name"),
            ("job", "Job Title"),
            ("email", "Email"),
            ("greet", "Greeting"),
        ]

        for field, label in basic_fields:
            self.inputs[field] = QLineEdit()
            left_group_layout.addRow(label + ":", self.inputs[field])

        left_group.setLayout(left_group_layout)
        left_layout.addWidget(left_group)

        # Контакты тоже в левой колонке
        contact_group = QGroupBox("Contact Information")
        contact_layout = QFormLayout()
        contact_layout.setSpacing(8)

        contact_fields = [
            ("work_number", "Work Phone"),
            ("personal_number", "Personal Phone"),
            ("social_number", "Social Number"),
            ("cut_number", "Cut Number"),
        ]

        for field, label in contact_fields:
            self.inputs[field] = QLineEdit()
            contact_layout.addRow(label + ":", self.inputs[field])

        contact_group.setLayout(contact_layout)
        left_layout.addWidget(contact_group)
        left_layout.addStretch()

        # ПРАВАЯ КОЛОНКА - Настройки
        settings_group = QGroupBox("Settings")
        settings_layout = QFormLayout()
        settings_layout.setSpacing(8)

        # Комбобоксы
        self.cb_hotel = QComboBox()
        self.cb_hotel.addItems(["", "Hyatt Regency", "Hyatt Place", "Both"])
        settings_layout.addRow("Hotel:", self.cb_hotel)

        self.cb_language = QComboBox()
        self.cb_language.addItems(["", "Russian", "English"])
        settings_layout.addRow("Language:", self.cb_language)

        self.cb_type = QComboBox()
        self.cb_type.addItems(["", "Full", "Cut"])
        settings_layout.addRow("Type:", self.cb_type)

        settings_group.setLayout(settings_layout)
        right_layout.addWidget(settings_group)

        # URL поля
        url_group = QGroupBox("URLs")
        url_layout = QFormLayout()
        url_layout.setSpacing(8)

        url_fields = [
            ("banner_path", "Banner Path"),
            ("banner_url", "Banner URL"),
            ("site_url", "Site URL"),
        ]

        for field, label in url_fields:
            self.inputs[field] = QLineEdit()
            url_layout.addRow(label + ":", self.inputs[field])

        url_group.setLayout(url_layout)
        right_layout.addWidget(url_group)

        # Чекбоксы в правой колонке
        features_group = QGroupBox("Enable Features")
        features_layout = QGridLayout()
        features_layout.setSpacing(10)

        self.checkboxes = {}
        conf_fields = [
            ("conf_greet", "Enable Greeting"),
            ("conf_fname", "Enable First Name"),
            ("conf_job", "Enable Job"),
            ("conf_hotel", "Enable Hotel"),
            ("conf_phone_numbers", "Enable Phone Numbers"),
            ("conf_mail", "Enable Email"),
            ("conf_banner", "Enable Banner"),
            ("conf_site", "Enable Site"),
            ("conf_main_sig", "Enable Main Signature"),
        ]

        # Располагаем чекбоксы в 2 колонки для лучшего вида
        for i, (field, label) in enumerate(conf_fields):
            self.checkboxes[field] = QCheckBox(label)
            row = i // 2  # 2 колонки
            col = i % 2
            features_layout.addWidget(self.checkboxes[field], row, col)

        features_group.setLayout(features_layout)
        right_layout.addWidget(features_group)
        right_layout.addStretch()

        # Добавляем колонки в горизонтальный layout
        columns_layout.addWidget(left_column)
        columns_layout.addWidget(right_column)

        scroll_layout.addWidget(columns_widget)
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)

        # Кнопки
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")

        save_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)

        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

    def load_user_data(self):
        if not self.signature_id or not self.db:
            return

        data = self.db.get_user_by_id(self.signature_id)
        if not data or len(data) == 0:
            return

        row = data[0]

        # Заполняем текстовые поля
        for field, widget in self.inputs.items():
            # Находим индекс поля
            field_indices = {
                'signature_id': 0, 'global_id': 1, 'signature_name': 2, 'first_name': 3,
                'last_name': 4, 'job': 5, 'email': 6, 'greet': 7,
                'work_number': 8, 'personal_number': 9, 'social_number': 10,
                'cut_number': 11, 'banner_path': 15, 'banner_url': 16,
                'site_url': 17
            }

            if field in field_indices:
                idx = field_indices[field]
                if idx < len(row) and row[idx] is not None:
                    widget.setText(str(row[idx]))

        # Комбобоксы
        if len(row) > 12 and row[12] is not None:
            try:
                self.cb_hotel.setCurrentIndex(int(row[12]))
            except:
                pass

        if len(row) > 13 and row[13] is not None:
            try:
                self.cb_language.setCurrentIndex(int(row[13]))
            except:
                pass

        if len(row) > 14 and row[14] is not None:
            try:
                self.cb_type.setCurrentIndex(int(row[14]))
            except:
                pass

        # Чекбоксы
        checkbox_indices = {
            'conf_greet': 18, 'conf_fname': 19, 'conf_job': 20,
            'conf_hotel': 21, 'conf_phone_numbers': 22, 'conf_mail': 23,
            'conf_banner': 24, 'conf_site': 25, 'conf_main_sig': 26
        }

        for field, checkbox in self.checkboxes.items():
            if field in checkbox_indices:
                idx = checkbox_indices[field]
                if idx < len(row) and row[idx] is not None:
                    checkbox.setChecked(row[idx] == 1)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())