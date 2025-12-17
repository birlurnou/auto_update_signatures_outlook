import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

class DatabaseManager:
    def __init__(self):
        # read ini file
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

    def execute_query(self, query, params=None, fetch_one=False, fetch_all=False, commit=False):
        conn = None
        cursor = None
        try:
            conn = self.get_connection()
            if not conn:
                return None
            cursor = conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            if commit:
                conn.commit()
                return True
            if fetch_one:
                return cursor.fetchone()
            elif fetch_all:
                return cursor.fetchall()
            return None
        except pyodbc.Error as e:
            return None
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def get_user_data(self, user_id):
        query = 'SELECT * FROM users WHERE user_id = ?'
        return self.execute_query(query, [user_id], fetch_one=True)

    def insert_signature(self, user_data):
        query = '''
        INSERT INTO ... (..., ..., ...)
        VALUES (?, ?, ?)
        '''
        return self.execute_query(query, user_data, commit=True)

    def update_signature(self, user_data):
        query = '''
        UPDATE ...
        SET ...=?, ...=?
        WHERE ...=?
        '''
        return self.execute_query(query, user_data, commit=True)

    def delete_signature(self, user_data):
        query = f'DELETE FROM ... WHERE ...=?'
        return self.execute_query(query, user_data, commit=True)

class SignatureApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Signature`s manager")
        window_width = 1000
        window_height = 745
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Темная тема и синий цвет по умолчанию
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Флаг текущей темы
        self.is_dark_theme = True

        # Хранилище подписей
        self.signatures = {}
        self.current_signature = None

        # Создаём главную вкладку напрямую
        self.create_main_tab()

    def update_treeview_style(self):
        """Обновляет стили Treeview в зависимости от темы"""
        style = ttk.Style()
        
        if self.is_dark_theme:
            # Темная тема
            style.configure("Treeview",
                            background="#2a2d2e",
                            foreground="white",
                            fieldbackground="#2a2d2e",
                            borderwidth=1,
                            relief="solid",
                            font=('Arial', 11))
            style.configure("Treeview.Heading",
                            background="#3b3b3b",
                            foreground="white",
                            relief="raised",
                            font=('Arial', 12, 'bold'))
            style.map('Treeview', background=[('selected', '#1f6aa5')])
            style.map("Treeview.Heading",
                      background=[('active', '#3b3b3b')],
                      foreground=[('active', '#AAAAAA')])
        else:
            # Светлая тема
            style.configure("Treeview",
                            background="white",
                            foreground="black",
                            fieldbackground="white",
                            borderwidth=1,
                            relief="solid",
                            font=('Arial', 11))
            style.configure("Treeview.Heading",
                            background="#e0e0e0",
                            foreground="black",
                            relief="raised",
                            font=('Arial', 12, 'bold'))
            style.map('Treeview', background=[('selected', '#3b8ed0')])
            style.map("Treeview.Heading",
                      background=[('active', '#d0d0d0')],
                      foreground=[('active', '#666666')])

    def create_main_tab(self):
        """Создаёт основную вкладку с таблицей подписей и кнопками действий"""
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Делим вкладку на левую и правую части
        left_frame = ctk.CTkFrame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        right_frame = ctk.CTkFrame(main_frame, width=150)
        right_frame.pack(side="right", fill="y", padx=5, pady=5)

        # Панель поиска (вместо заголовка)
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.pack(fill="x", pady=5, padx=5)

        # Поле для ввода поиска
        self.search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="Введите текст для поиска...",
            width=300,
            height=35,
            corner_radius=6,
            font=("Arial", 12)
        )
        self.search_entry.pack(side="left", padx=(0, 10))

        # Кнопка "Найти"
        self.search_button = ctk.CTkButton(
            search_frame,
            text="Найти",
            command=self.on_search_click,
            width=80,
            height=35,
            corner_radius=6,
            fg_color="#3b8ed0",
            hover_color="#1f6aa5",
            font=("Arial", 12, "bold")
        )
        self.search_button.pack(side="left")

        # Создаем таблицу (Treeview)
        table_frame = ctk.CTkFrame(left_frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Создаем Treeview с кастомными стилями
        style = ttk.Style()
        style.theme_use("default")
        
        # Инициализируем стили для темной темы
        self.update_treeview_style()

        # Создаем Treeview
        self.tree = ttk.Treeview(table_frame, columns=("global_id", "first_name", "last_name", "department"),
                                 show="headings", height=15, style="Treeview")

        # Настраиваем колонки
        self.tree.heading("global_id", text="Global ID")
        self.tree.heading("first_name", text="First Name")
        self.tree.heading("last_name", text="Last Name")
        self.tree.heading("department", text="Department")

        self.tree.column("global_id", width=150, anchor="center", stretch=False)
        self.tree.column("first_name", width=150, anchor="w", stretch=False)
        self.tree.column("last_name", width=150, anchor="w", stretch=False)
        self.tree.column("department", width=550, anchor="w", stretch=False)

        # Делаем заголовки не кликабельными
        for column in ("global_id", "first_name", "last_name", "department"):
            self.tree.heading(column, command=lambda: None)  # Пустая функция

        # Добавляем скроллбар для таблицы
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Упаковываем таблицу и скроллбар
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Добавляем примеры данных в таблицу
        self.add_sample_data()

        # Обработчик выбора строки в таблице
        self.tree.bind("<<TreeviewSelect>>", self.on_table_select)

        # Кнопки действий
        ctk.CTkLabel(right_frame, text="Действия", font=("Arial", 14, "bold")).pack(pady=10)

        self.btn_create = ctk.CTkButton(
            right_frame, text="Создать", command=self.create_signature, height=40, width=140,
            fg_color="#3b8ed0", hover_color="#1f6aa5", text_color="white",
            font=("Arial", 14)
        )
        self.btn_create.pack(pady=5)

        self.btn_delete = ctk.CTkButton(
            right_frame, text="Удалить", command=self.delete_signature, height=40, width=140,
            fg_color="#a0a0a0",  # Серый цвет по умолчанию (неактивное состояние)
            hover_color="#a0a0a0",  # Тот же цвет при наведении в неактивном состоянии
            text_color="white",
            text_color_disabled="white",
            state="disabled",
            font=("Arial", 14)
        )
        self.btn_delete.pack(pady=5)

        # Кнопка переключения темы
        self.theme_button = ctk.CTkButton(
            right_frame, 
            text="Светлая тема", 
            command=self.toggle_theme, 
            height=40, 
            width=140,
            fg_color="#FFFFFF",  # Белый цвет
            hover_color="#E0E0E0",  # Светло-серый при наведении
            text_color="#000000",  # Черный текст
            font=("Arial", 14)
        )
        self.theme_button.pack(pady=5)

    def toggle_theme(self):
        """Переключает между светлой и темной темой"""
        if self.is_dark_theme:
            # Переключаем на светлую тему
            ctk.set_appearance_mode("light")
            self.theme_button.configure(
                text="Темная тема",
                fg_color="#000000",  # Черный цвет
                hover_color="#333333",  # Темно-серый при наведении
                text_color="#FFFFFF"  # Белый текст
            )
            self.is_dark_theme = False
        else:
            # Переключаем на темную тему
            ctk.set_appearance_mode("dark")
            self.theme_button.configure(
                text="Светлая тема",
                fg_color="#FFFFFF",  # Белый цвет
                hover_color="#E0E0E0",  # Светло-серый при наведении
                text_color="#000000"  # Черный текст
            )
            self.is_dark_theme = True
        
        # Обновляем стили Treeview
        self.update_treeview_style()
        
        # Принудительно обновляем отображение таблицы
        self.tree.update()

    def on_search_click(self):
        """Обработчик нажатия кнопки 'Найти'"""
        search_text = self.search_entry.get()
        print(f"Поиск: {search_text}")  # Вывод в консоль
        # В разработке

    def add_sample_data(self):
        """Добавляет примеры данных в таблицу"""
        sample_data = [
            ("GLB001", "John", "Smith", "IT Department"),
            ("GLB002", "Anna", "Johnson", "Sales Department"),
            ("GLB003", "Mike", "Williams", "HR Department"),
            ("GLB004", "Sarah", "Brown", "Marketing"),
            ("GLB005", "David", "Miller", "Finance"),
            ("GLB006", "Emma", "Davis", "Operations"),
            ("GLB007", "James", "Wilson", "IT Department"),
            ("GLB008", "Lisa", "Moore", "Sales Department"),
            ("GLB009", "Robert", "Taylor", "HR Department"),
            ("GLB010", "Maria", "Anderson", "Marketing")
        ]

        for data in sample_data:
            self.tree.insert("", "end", values=data)

    def on_table_select(self, event):
        """Обработчик выбора строки в таблице"""
        selected_items = self.tree.selection()
        if selected_items:
            self.btn_delete.configure(
                state="normal",
                fg_color="#d9534f",  # Красный цвет когда активна
                hover_color="#c9302c"  # Красный ховер когда активна
            )
            # Получаем данные выбранной строки
            item = selected_items[0]
            values = self.tree.item(item, "values")
            self.current_signature = {
                "global_id": values[0],
                "first_name": values[1],
                "last_name": values[2],
                "department": values[3]
            }
        else:
            self.btn_delete.configure(
                state="disabled",
                fg_color="#a0a0a0",  # Серый цвет когда неактивна
                hover_color="#a0a0a0"  # Серый ховер когда неактивна
            )
            self.current_signature = None

    def create_signature(self):
        """Создаёт новую подпись"""
        # Добавляем тестовые данные
        new_id = f"GLB{len(self.tree.get_children()) + 1:03d}"
        new_data = (new_id, "New", "User", "New Department")
        self.tree.insert("", "end", values=new_data)

    def delete_signature(self):
        """Удаляет выбранную подпись"""
        if not self.current_signature:
            messagebox.showwarning("Ошибка", "Выберите подпись из таблицы")
            return

        if messagebox.askyesno("Подтверждение",
                               f"Удалить подпись '{self.current_signature['first_name']} {self.current_signature['last_name']}'?"):
            selected_items = self.tree.selection()
            for item in selected_items:
                self.tree.delete(item)

            self.current_signature = None
            self.btn_delete.configure(
                state="disabled",
                fg_color="#a0a0a0",  # Серый цвет после удаления
                hover_color="#a0a0a0"  # Серый ховер после удаления
            )

if __name__ == "__main__":
    app = SignatureApp()
    app.mainloop()