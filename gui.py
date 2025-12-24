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
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.doubleClicked.connect(self.on_double_click)

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
        self.edit_btn = QPushButton("Edit")
        self.delete_btn = QPushButton("Delete")

        buttons = [self.create_btn, self.edit_btn, self.delete_btn]
        for btn in buttons:
            btn.setMinimumHeight(35)
            btn.setMinimumWidth(150)
            right_layout.addWidget(btn)

        right_layout.addStretch()

        # Подключаем кнопки
        self.create_btn.clicked.connect(self.on_create)
        self.edit_btn.clicked.connect(self.on_edit)
        self.delete_btn.clicked.connect(self.on_delete)

        layout.addWidget(left_widget, 3)
        layout.addWidget(right_widget, 1)

    def load_data(self, data=None):
        if data is None:
            data = self.db.get_all_users()

        self.table.setRowCount(0)
        if data:
            self.table.setRowCount(len(data))
            for row_idx, row in enumerate(data):
                for col_idx, value in enumerate(row):
                    item = QTableWidgetItem(str(value) if value else "")
                    self.table.setItem(row_idx, col_idx, item)

        self.table.resizeColumnsToContents()

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
            dialog = SimpleEditDialog(global_id, self.db, self)
            if dialog.exec_():
                self.on_search()
        else:
            QMessageBox.warning(self, "Warning", "Select a signature first")

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
        global_id = self.table.item(row, 0).text()
        dialog = SimpleEditDialog(global_id, self.db, self)
        if dialog.exec_():
            self.on_search()


class SimpleEditDialog(QDialog):
    def __init__(self, global_id=None, db=None, parent=None):
        super().__init__(parent)
        self.global_id = global_id
        self.db = db
        self.initUI()
        if global_id:
            self.load_user_data()

    def initUI(self):
        title = "Edit Signature" if self.global_id else "Create Signature"
        self.setWindowTitle(title)
        self.setMinimumWidth(500)

        layout = QVBoxLayout()

        # Поля с правильными названиями
        fields = [
            ("global_id", "Global ID", 0),
            ("signature_name", "Signature Name", 1),
            ("first_name", "First Name", 2),
            ("last_name", "Last Name", 3),
            ("job", "Job Title", 4),
            ("email", "Email", 5),
            ("greet", "Greeting", 6),
            ("work_number", "Work Phone", 7),
            ("personal_number", "Personal Phone", 8),
            ("social_number", "Social Number", 9),
            ("cut_number", "Cut Number", 10),
        ]

        self.inputs = {}
        for field, label, idx in fields:
            self.inputs[field] = QLineEdit()
            layout.addWidget(QLabel(label))
            layout.addWidget(self.inputs[field])

        # Комбобоксы
        layout.addWidget(QLabel("Hotel (1=Hyatt Regency, 2=Hyatt Place, 3=Both)"))
        self.cb_hotel = QComboBox()
        self.cb_hotel.addItems(["", "Hyatt Regency", "Hyatt Place", "Both"])
        layout.addWidget(self.cb_hotel)

        layout.addWidget(QLabel("Language (1=ru, 2=en)"))
        self.cb_language = QComboBox()
        self.cb_language.addItems(["", "Russian", "English"])
        layout.addWidget(self.cb_language)

        layout.addWidget(QLabel("Type (1=full, 2=cut)"))
        self.cb_type = QComboBox()
        self.cb_type.addItems(["", "Full", "Cut"])
        layout.addWidget(self.cb_type)

        # URL поля
        url_fields = [
            ("banner_path", "Banner Path", 14),
            ("banner_url", "Banner URL", 15),
            ("site_url", "Site URL", 16),
        ]

        for field, label, idx in url_fields:
            self.inputs[field] = QLineEdit()
            layout.addWidget(QLabel(label))
            layout.addWidget(self.inputs[field])

        # Чекбоксы
        self.checkboxes = {}
        conf_fields = [
            ("conf_greet", "Enable Greeting", 17),
            ("conf_fname", "Enable First Name", 18),
            ("conf_job", "Enable Job", 19),
            ("conf_hotel", "Enable Hotel", 20),
            ("conf_phone_numbers", "Enable Phone Numbers", 21),
            ("conf_mail", "Enable Email", 22),
            ("conf_banner", "Enable Banner", 23),
            ("conf_site", "Enable Site", 24),
            ("conf_main_sig", "Enable Main Signature", 25),
        ]

        for field, label, idx in conf_fields:
            self.checkboxes[field] = QCheckBox(label)
            layout.addWidget(self.checkboxes[field])

        # Кнопки
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")

        save_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def load_user_data(self):
        if not self.global_id or not self.db:
            return

        data = self.db.get_user_by_id(self.global_id)
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
                    # checkbox.setChecked(bool(row[idx]))
                    checkbox.setChecked(row[idx] == 1)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())