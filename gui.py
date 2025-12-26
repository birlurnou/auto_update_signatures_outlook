import os
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import configparser
import pyodbc
import tempfile
import webbrowser


def create_email_signature(first_name, last_name, job, email, greet,
                           work_number, personal_number, social_number, cut_number,
                           cb_hotel, cb_language, cb_type,
                           banner_path, banner_url, site_url,
                           conf_greet,
                           conf_fname,
                           conf_job,
                           conf_hotel,
                           conf_phone_numbers,
                           conf_mail,
                           conf_banner,
                           conf_site,
                           ):
    # read ini file
    config = configparser.ConfigParser()
    config.read('config.ini')
    # [hotel]
    color_rg = config['hotel']['color_rg']
    color_ze = config['hotel']['color_ze']
    address_rg_ru = config['hotel']['address_rg_ru']
    address_ze_ru = config['hotel']['address_ze_ru']
    address_rg_en = config['hotel']['address_rg_en']
    address_ze_en = config['hotel']['address_ze_en']
    hotel_name_rg_ru = config['hotel']['hotel_name_rg_ru']
    hotel_name_ze_ru = config['hotel']['hotel_name_ze_ru']
    hotel_name_rg_en = config['hotel']['hotel_name_rg_en']
    hotel_name_ze_en = config['hotel']['hotel_name_ze_en']

    # color and config from hotel
    hotel_config = None
    if cb_language == 1:
        hotel_config = {
            1: {
                'color': f'{color_rg}',
                'address': f'{address_rg_ru}',
                'hotel_name': [f'{hotel_name_rg_ru}']
            },
            2: {
                'color': f'{color_ze}',
                'address': f'{address_ze_ru}',
                'hotel_name': [f'{hotel_name_ze_ru}']
            },
            3: {
                'color': f'{color_rg}',
                'address': f'{address_rg_ru}',
                'hotel_name': [f'{hotel_name_rg_ru}', f'{hotel_name_ze_ru}']
            }
        }
    elif cb_language == 2:
        hotel_config = {
            1: {
                'color': f'{color_rg}',
                'address': f'{address_rg_en}',
                'hotel_name': [f'{hotel_name_rg_en}']
            },
            2: {
                'color': f'{color_ze}',
                'address': f'{address_ze_en}',
                'hotel_name': [f'{hotel_name_ze_en}']
            },
            3: {
                'color': f'{color_rg}',
                'address': f'{address_rg_en}',
                'hotel_name': [f'{hotel_name_rg_en}', f'{hotel_name_ze_en}']
            }
        }
    if not hotel_config:
        return None
    config_hotel = hotel_config.get(cb_hotel, None)

    # full name
    full_name = f'{first_name} {last_name}'.upper()

    # [phones]
    work_tag = config['phones']['work_tag']
    pers_tag = config['phones']['pers_tag']
    soc_tag = config['phones']['soc_tag']

    phones_html = ''

    # add phones
    if conf_phone_numbers == 1:
        # work phones
        if work_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]};
mso-bidi-font-weight:bold'> {work_tag}&nbsp;&nbsp;</span><span style='font-size:10.0pt;
font-family:"Arial",sans-serif;color:black'> {work_number}&nbsp;&nbsp; <o:p></o:p></span></p>'''

        # mobile + wa
        if personal_number and social_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]}; mso-bidi-font-weight:bold'> {pers_tag}&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {personal_number}&nbsp;&nbsp;</span>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]}; mso-bidi-font-weight:bold'> {soc_tag}&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {social_number}&nbsp;&nbsp;</span>
<o:p></o:p>
</p>'''
        elif personal_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]};
mso-bidi-font-weight:bold'>{pers_tag}&nbsp;</span><span style='font-size:10.0pt;font-family:
"Arial",sans-serif;color:black'> {personal_number}&nbsp;&nbsp; <o:p></o:p></span></p>'''
        elif social_number and not work_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]};
mso-bidi-font-weight:bold'>{soc_tag}</span><span style='font-size:10.0pt;font-family:
"Arial",sans-serif;color:black'> {social_number}&nbsp;&nbsp; <o:p></o:p></span></p>'''
        elif social_number != '':
            phones_html = f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]}; mso-bidi-font-weight:bold'> {work_tag}&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {work_number}&nbsp;&nbsp; <o:p></o:p></span>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]}; mso-bidi-font-weight:bold'> {soc_tag}&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {social_number}&nbsp;&nbsp;</span>
<o:p></o:p>
</p>'''

    # greeting
    if greet and conf_greet == 1:
        greeting = f'''
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Arial",sans-serif;
color:black'>{greet},<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Arial",sans-serif;
color:black'><o:p>&nbsp;</o:p></span></p>'''

    else:
        greeting = ''

    # f name
    full_name_html = ''
    if conf_fname == 1:
        full_name_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:10.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:{config_hotel["color"]};text-transform:
uppercase'>{full_name}<o:p></o:p></span></b></p>'''

    # job
    job_html = ''
    space_after_job = ''
    if conf_job == 1:
        job_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><span style='font-size:10.0pt;mso-bidi-font-size:
9.0pt;line-height:120%;font-family:"Arial",sans-serif;mso-bidi-font-weight:
bold'>{job.capitalize()}<o:p></o:p></span></p>'''

        # space_after_job
        space_after_job = '''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:
Calibri'><o:p>&nbsp;</o:p></span></p>'''

    # hotel and address
    hotel_and_address = ''
    space_after_address = ''
    if conf_hotel == 1:
        for item in config_hotel["hotel_name"]:
            hotel_and_address += f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:"Arial",sans-serif;
color:{config_hotel["color"]}'>{item.replace("<br>", "<o:p></o:p></span></p><p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:\"Arial\",sans-serif; color:{config[\"color\"]}'>")}<o:p></o:p></span></p>'''
        hotel_and_address += f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:"Arial",sans-serif'>
{config_hotel["address"]}<o:p></o:p></span></p>'''

        # space_after_address
        space_after_address = '''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>'''

    # email
    email_html = ''
    if conf_mail == 1:
        email_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]}'> E&nbsp;&nbsp;&nbsp;</span><span
lang=EN-US style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-ansi-language:
EN-US'><a href="mailto:{email}">{email}</a></span><span
style='font-size:10.0pt;font-family:"Arial",sans-serif'><o:p></o:p></span></p>'''

    # banner
    banner_html = ''
    if conf_banner == 1 and banner_path:
        if banner_path[-3:] == 'png' or banner_path[-3:] == 'jpg':
            try:
                with open(banner_path, 'rb'):
                    banner_html = f'''
                    <p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
                    line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
                    line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>

                    <p class=MsoNormal><a href="{banner_url}">
                    <img border=0 width=779 height=136 src="{banner_path}" style="border:none;">
                    </a><o:p></o:p></p>'''

            except FileNotFoundError:
                if banner_path[:-3] == 'png':
                    banner_path = banner_path[:-3] + 'jpg'
                else:
                    banner_path = banner_path[:-3] + 'png'

                try:

                    with open(banner_path, 'rb'):
                        banner_html = f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>

<p class=MsoNormal><a href="{banner_url}">
<img border=0 width=779 height=136 src="{banner_path}" style="border:none;">
</a><o:p></o:p></p>'''
                except:
                    banner_html = ''

    # site
    site_html = ''
    if conf_site == 1:
        site_html = f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>

<p class='MsoNormal' style='text-align:justify; text-justify:inter-ideograph; font-size:10.0pt; font-family:Arial, sans-serif; color:{config_hotel["color"]};'>
    <a href='{site_url}' style='color: inherit; text-decoration: none;'>{site_url}</a>
</p>'''

    # full html
    if cb_type != 1:
        space_after_job = ''
        hotel_and_address = ''
        space_after_address = ''
        phones_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config_hotel["color"]};
mso-bidi-font-weight:bold'></span><span style='font-size:10.0pt;
font-family:"Arial",sans-serif;color:black'>{cut_number if cut_number else ''}&nbsp;&nbsp; <o:p></o:p></span></p>'''
        email_html = ''
    html_signature = f'''<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1251">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">
<!--[if !mso]>
<style>
v\\:* {{behavior:url(#default#VML);}}
o\\:* {{behavior:url(#default#VML);}}
w\\:* {{behavior:url(#default#VML);}}
.shape {{behavior:url(#default#VML);}}
</style>
<![endif]-->
<style>
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}}
</style>
</head>
<body lang=RU link="#0563C1" vlink="#954F72" style='tab-interval:35.4pt'>
<div class=WordSection1>

{greeting}

{full_name_html}

{job_html}

{space_after_job}

{hotel_and_address}

{space_after_address}

{phones_html}

{email_html}

{banner_html}

{site_html}

<p class=MsoNormal><o:p>&nbsp;</o:p></p>
</div>
</body>
</html>'''

    return html_signature

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

    def execute_update(self, query, params=None):
        """Выполняет UPDATE запрос"""
        conn = self.get_connection()
        if not conn:
            return False
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
            return True
        except pyodbc.Error as e:
            print(f'Update error: {e}')
            conn.rollback()
            return False
        finally:
            cursor.close()
            conn.close()

    def execute_insert(self, query, params=None):
        """Выполняет INSERT запрос"""
        conn = self.get_connection()
        if not conn:
            return False
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
            return True
        except pyodbc.Error as e:
            print(f'Insert error: {e}')
            conn.rollback()
            return False
        finally:
            cursor.close()
            conn.close()

    def execute_delete(self, query, params=None):
        """Выполняет DELETE запрос"""
        conn = self.get_connection()
        if not conn:
            return False
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
            return True
        except pyodbc.Error as e:
            print(f'Delete error: {e}')
            conn.rollback()
            return False
        finally:
            cursor.close()
            conn.close()

    def get_all_users(self):
        query = 'SELECT signature_id, global_id, signature_name, first_name, last_name, email, cb_hotel, cb_type, conf_main_sig, conf_greet, conf_banner, conf_site FROM signatures ORDER BY signature_id ASC'
        return self.execute_query(query, fetch_all=True)

    def search_users(self, search_term):
        query = '''
        SELECT signature_id, global_id, signature_name, first_name, last_name, email, 
               cb_hotel, cb_type, conf_main_sig, conf_greet, conf_banner, conf_site 
        FROM signatures 
        WHERE global_id LIKE ? OR signature_name LIKE ? OR first_name LIKE ? 
        OR last_name LIKE ? OR email LIKE ?
        '''
        pattern = f'%{search_term}%'
        return self.execute_query(query, [pattern] * 5, fetch_all=True)

    def get_user_by_id(self, signature_id):
        query = 'SELECT * FROM signatures WHERE signature_id = ?'
        result = self.execute_query(query, [signature_id], fetch_all=True)
        return result

    def insert_user(self, data):
        """Вставляет новую запись в таблицу signatures"""
        query = '''
        INSERT INTO signatures (
            global_id, signature_name, first_name, last_name, job, email, greet,
            work_number, personal_number, social_number, cut_number,
            cb_hotel, cb_language, cb_type,
            banner_path, banner_url, site_url,
            conf_greet, conf_fname, conf_job, conf_hotel, conf_phone_numbers,
            conf_mail, conf_banner, conf_site, conf_main_sig
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        '''
        return self.execute_insert(query, data)

    def update_user(self, signature_id, data):
        """Обновляет существующую запись в таблице signatures"""
        query = '''
        UPDATE signatures SET
            global_id = ?, signature_name = ?, first_name = ?, last_name = ?, job = ?, email = ?, greet = ?,
            work_number = ?, personal_number = ?, social_number = ?, cut_number = ?,
            cb_hotel = ?, cb_language = ?, cb_type = ?,
            banner_path = ?, banner_url = ?, site_url = ?,
            conf_greet = ?, conf_fname = ?, conf_job = ?, conf_hotel = ?, conf_phone_numbers = ?,
            conf_mail = ?, conf_banner = ?, conf_site = ?, conf_main_sig = ?
        WHERE signature_id = ?
        '''
        # Добавляем signature_id в конец списка параметров
        params = data + [signature_id]
        return self.execute_update(query, params)

    def delete_user(self, signature_id):
        """Удаляет запись из таблицы signatures"""
        query = 'DELETE FROM signatures WHERE signature_id = ?'
        return self.execute_delete(query, [signature_id])

    def update_global_settings(self, banner_path_rg, banner_url_rg, site_url_rg,
                               banner_path_ze, banner_url_ze, site_url_ze):
        """Обновляет глобальные настройки для всех подписей"""
        success = True

        # Первый UPDATE для Regency (cb_hotel = 1 или 3)
        query_rg = '''
        UPDATE signatures 
        SET banner_path = ?, banner_url = ?, site_url = ? 
        WHERE cb_hotel = 1 OR cb_hotel = 3
        '''
        success_rg = self.execute_update(query_rg, [banner_path_rg, banner_url_rg, site_url_rg])

        # Второй UPDATE для Place (cb_hotel = 2)
        query_ze = '''
        UPDATE signatures 
        SET banner_path = ?, banner_url = ?, site_url = ? 
        WHERE cb_hotel = 2
        '''
        success_ze = self.execute_update(query_ze, [banner_path_ze, banner_url_ze, site_url_ze])

        return success_rg and success_ze


class GlobalSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = DatabaseManager()
        self.setWindowIcon(QIcon('icon.ico'))
        self.initUI()
        font = self.font()
        font.setPointSize(9)  # Увеличить размер
        self.setFont(font)
        self.load_settings()

    def initUI(self):
        self.setWindowTitle('Global Settings')
        self.setMinimumWidth(1000)
        self.setMinimumHeight(200)

        # Основной layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # Две колонки
        columns_widget = QWidget()
        columns_layout = QHBoxLayout(columns_widget)
        columns_layout.setSpacing(20)

        # Первая колонка - Regency
        regency_group = QGroupBox("Regency")
        regency_layout = QFormLayout()
        regency_layout.setSpacing(10)

        # Поля для Regency
        self.banner_path_rg = QLineEdit()
        regency_layout.addRow("Banner Path (RG):", self.banner_path_rg)

        self.banner_url_rg = QLineEdit()
        regency_layout.addRow("Banner URL (RG):", self.banner_url_rg)

        self.site_url_rg = QLineEdit()
        regency_layout.addRow("Site URL (RG):", self.site_url_rg)

        regency_group.setLayout(regency_layout)

        # Вторая колонка - Place
        place_group = QGroupBox("Place")
        place_layout = QFormLayout()
        place_layout.setSpacing(10)

        # Поля для Place
        self.banner_path_ze = QLineEdit()
        place_layout.addRow("Banner Path (ZE):", self.banner_path_ze)

        self.banner_url_ze = QLineEdit()
        place_layout.addRow("Banner URL (ZE):", self.banner_url_ze)

        self.site_url_ze = QLineEdit()
        place_layout.addRow("Site URL (ZE):", self.site_url_ze)

        place_group.setLayout(place_layout)

        # Добавляем колонки
        columns_layout.addWidget(regency_group)
        columns_layout.addWidget(place_group)

        main_layout.addWidget(columns_widget)

        # Кнопки
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        # preview_btn = QPushButton("Preview")
        cancel_btn = QPushButton("Cancel")

        save_btn.clicked.connect(self.on_save)
        # preview_btn.clicked.connect(self.on_preview)
        cancel_btn.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        # button_layout.addWidget(preview_btn)
        button_layout.addWidget(cancel_btn)

        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

    def load_settings(self):
        """Загружает настройки из config.ini"""
        config = configparser.ConfigParser()
        config.read('config.ini')

        # Загружаем значения из config.ini
        if 'global' in config:
            self.banner_path_rg.setText(config['global'].get('banner_path_rg', ''))
            self.banner_url_rg.setText(config['global'].get('banner_url_rg', ''))
            self.site_url_rg.setText(config['global'].get('site_url_rg', ''))
            self.banner_path_ze.setText(config['global'].get('banner_path_ze', ''))
            self.banner_url_ze.setText(config['global'].get('banner_url_ze', ''))
            self.site_url_ze.setText(config['global'].get('site_url_ze', ''))

    def on_save(self):
        """Обработчик нажатия кнопки Save"""
        # Получаем значения из полей
        banner_path_rg = self.banner_path_rg.text().strip()
        banner_url_rg = self.banner_url_rg.text().strip()
        site_url_rg = self.site_url_rg.text().strip()
        banner_path_ze = self.banner_path_ze.text().strip()
        banner_url_ze = self.banner_url_ze.text().strip()
        site_url_ze = self.site_url_ze.text().strip()

        # Сохраняем в config.ini
        config = configparser.ConfigParser()
        config.read('config.ini')

        if 'global' not in config:
            config['global'] = {}

        config['global']['banner_path_rg'] = banner_path_rg
        config['global']['banner_url_rg'] = banner_url_rg
        config['global']['site_url_rg'] = site_url_rg
        config['global']['banner_path_ze'] = banner_path_ze
        config['global']['banner_url_ze'] = banner_url_ze
        config['global']['site_url_ze'] = site_url_ze

        with open('config.ini', 'w') as configfile:
            config.write(configfile)

        # Выполняем UPDATE запросы в базу данных
        success = self.db.update_global_settings(
            banner_path_rg, banner_url_rg, site_url_rg,
            banner_path_ze, banner_url_ze, site_url_ze
        )

        if success:
            # QMessageBox.information(self, "Success", "Global settings saved successfully!")
            self.accept()
        else:
            QMessageBox.warning(self, "Error", "Failed to save settings to database!")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.setWindowIcon(QIcon('icon.ico'))
        self.initUI()
        self.load_data()
        # Изначально деактивируем кнопки, требующие выбора
        self.update_button_states()
        self.showMaximized()

    def initUI(self):
        self.setWindowTitle('Signature Manager')
        self.setGeometry(100, 100, 1000, 600)

        # Увеличить шрифт для всего главного окна
        font = self.font()
        font.setPointSize(9)
        self.setFont(font)

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
        self.table.setColumnCount(12)
        self.table.setHorizontalHeaderLabels(['ID', 'Global ID', 'Sig name', 'First name', 'Last name', 'Email', 'Hotel'
                                                 , 'Type', 'Main', 'Greeting', 'Banner', 'Site'])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.doubleClicked.connect(self.on_double_click)

        # ПОДКЛЮЧАЕМ СИГНАЛ selectionChanged - ВАЖНО!
        self.table.selectionModel().selectionChanged.connect(self.on_selection_changed)

        # Стиль таблицы
        palette = self.table.palette()
        palette.setColor(palette.Highlight, QColor(0, 120, 215))
        palette.setColor(palette.HighlightedText, Qt.white)
        self.table.setPalette(palette)

        left_layout.addLayout(search_layout)
        left_layout.addWidget(self.table)

        # Устанавливаем фильтр событий для таблицы
        self.table.viewport().installEventFilter(self)

        # Правая часть с кнопками
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setAlignment(Qt.AlignTop)
        right_layout.setSpacing(10)

        self.create_btn = QPushButton("Create")
        self.copy_btn = QPushButton("Copy")
        self.edit_btn = QPushButton("Edit")
        self.more_settings_btn = QPushButton("Global settings")
        self.delete_btn = QPushButton("Delete")

        buttons = [self.create_btn, self.copy_btn, self.edit_btn, self.more_settings_btn, self.delete_btn]
        for btn in buttons:
            btn.setMinimumHeight(35)
            btn.setMinimumWidth(150)
            right_layout.addWidget(btn)

        right_layout.addStretch()

        # Подключаем кнопки
        self.create_btn.clicked.connect(self.on_create)
        self.copy_btn.clicked.connect(self.on_copy)
        self.edit_btn.clicked.connect(self.on_edit)
        self.more_settings_btn.clicked.connect(self.on_more_settings)
        self.delete_btn.clicked.connect(self.on_delete)

        layout.addWidget(left_widget, 3)
        layout.addWidget(right_widget, 1)

    # выравнивание
    def eventFilter(self, source, event):
        """Обработчик событий для снятия выделения"""
        if source is self.table.viewport() and event.type() == QEvent.MouseButtonPress:
            # Получаем позицию клика
            pos = event.pos()
            # Проверяем, есть ли элемент под курсором
            item = self.table.itemAt(pos)

            if item is None:
                # Клик на пустое место - снимаем выделение
                self.table.clearSelection()
                # Обновляем состояние кнопок
                self.update_button_states()

        return super().eventFilter(source, event)

    def load_data(self, data=None):
        if data is None:
            data = self.db.get_all_users()

        self.table.setRowCount(0)
        if data:
            self.table.setRowCount(len(data))
            for row_idx, row in enumerate(data):
                for col_idx, value in enumerate(row):
                    item_text = str(value) if value else ""

                    if col_idx == 6:
                        hotel_map = {1: "Regency", 2: "Place", 3: "Both"}
                        item_text = hotel_map.get(value, "")
                    elif col_idx == 7:
                        type_map = {1: "Full", 2: "Cut"}
                        item_text = type_map.get(value, "")
                    elif col_idx == 8:
                        main_map = {1: '✔️', 2: ''}
                        item_text = main_map.get(value, "")
                    elif col_idx == 9:
                        greet_map = {1: '✔️', 2: ''}
                        item_text = greet_map.get(value, "")
                    elif col_idx == 10:
                        banner_map = {1: '✔️', 2: ''}
                        item_text = banner_map.get(value, "")
                    elif col_idx == 11:
                        site_map = {1: '✔️', 2: ''}
                        item_text = site_map.get(value, "")

                    item = QTableWidgetItem(item_text)

                    center_columns = [0, 6, 7, 8, 9, 10, 11]  # ID, Hotel, Type, Main, Greeting, Banner, Site
                    if col_idx in center_columns:
                        item.setTextAlignment(Qt.AlignCenter)

                    if col_idx == 6:  # Столбец Hotel
                        hotel_value = value
                        if hotel_value == 1:  # Regency
                            item.setBackground(QColor("#441D61"))  # Фиолетовый для Regency
                            item.setForeground(QColor("#FFFFFF"))  # Белый текст для контраста
                        elif hotel_value == 2:  # Place
                            item.setBackground(QColor("#ED7D31"))  # Оранжевый для Place
                            item.setForeground(QColor("#FFFFFF"))  # Белый текст для контраста
                        elif hotel_value == 3:  # Both
                            item.setBackground(QColor("#7B2CBF"))  # Средний фиолетовый
                            item.setForeground(QColor("#FFFFFF"))  # Белый текст

                    self.table.setItem(row_idx, col_idx, item)

        self.table.resizeColumnsToContents()
        # Обновляем состояние кнопок после загрузки данных
        self.update_button_states()

    def update_button_states(self):
        """Обновляет состояние кнопок в зависимости от выбора в таблице"""
        # Проверяем, есть ли выделенные строки
        selected_rows = self.table.selectionModel().selectedRows()

        # Если есть выделенные строки
        has_valid_selection = len(selected_rows) > 0

        # Кнопки, которые требуют выбора элемента
        self.copy_btn.setEnabled(has_valid_selection)
        self.edit_btn.setEnabled(has_valid_selection)
        self.delete_btn.setEnabled(has_valid_selection)

        # Кнопка Create всегда активна
        self.create_btn.setEnabled(True)
        self.more_settings_btn.setEnabled(True)

    def on_selection_changed(self, selected, deselected):
        """Обработчик изменения выделения в таблице"""
        self.update_button_states()

    def on_search(self):
        term = self.search_input.text().strip()
        data = self.db.search_users(term) if term else self.db.get_all_users()
        self.load_data(data)

    def on_create(self):
        dialog = SimpleEditDialog(None, self.db, self, mode='create')
        if dialog.exec_():
            self.on_search()

    def on_edit(self):
        """Обработчик кнопки Edit"""
        selected_rows = self.table.selectionModel().selectedRows()
        if selected_rows:
            row = selected_rows[0].row()
            signature_id = self.table.item(row, 0).text()
            dialog = SimpleEditDialog(signature_id, self.db, self, mode='edit')
            if dialog.exec_():
                self.on_search()
        else:
            QMessageBox.warning(self, "Warning", "Select a signature first")

    def on_copy(self):
        """Обработчик кнопки Copy"""
        selected_rows = self.table.selectionModel().selectedRows()
        if selected_rows:
            row = selected_rows[0].row()
            signature_id = self.table.item(row, 0).text()
            dialog = SimpleEditDialog(signature_id, self.db, self, mode='copy')
            if dialog.exec_():
                self.on_search()
        else:
            QMessageBox.warning(self, "Warning", "Select a signature first")

    def on_more_settings(self):
        dialog = GlobalSettingsDialog(self)
        dialog.exec_()

    def on_delete(self):
        """Обработчик кнопки Delete"""
        selected_rows = self.table.selectionModel().selectedRows()
        if selected_rows:
            row = selected_rows[0].row()
            signature_id = self.table.item(row, 0).text()
            reply = QMessageBox.question(self, "Confirm", "Delete this signature?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                success = self.db.delete_user(signature_id)
                if success:
                    self.on_search()
                else:
                    QMessageBox.warning(self, "Error", "Failed to delete signature!")
        else:
            QMessageBox.warning(self, "Warning", "Select a signature first")

    def on_double_click(self, index):
        row = index.row()
        signature_id = self.table.item(row, 0).text()
        dialog = SimpleEditDialog(signature_id, self.db, self, mode='edit')
        if dialog.exec_():
            self.on_search()


class SimpleEditDialog(QDialog):
    def __init__(self, signature_id=None, db=None, parent=None, mode='edit'):
        super().__init__(parent)
        self.signature_id = signature_id
        self.db = db
        self.mode = mode  # 'create', 'edit', or 'copy'
        self.setWindowIcon(QIcon('icon.ico'))
        self.initUI()
        if signature_id and mode in ['edit', 'copy']:
            self.load_user_data()
            if mode == 'copy':
                ...
                # self.clear_global_id()  # Очищаем global_id для копирования
                # self.clear_signature_name() # Очищаем signature_name для копирования
        font = self.font()
        font.setPointSize(9)  # Увеличить размер
        self.setFont(font)

    def initUI(self):
        if self.mode == 'create':
            title = "Create Signature"
        elif self.mode == 'copy':
            title = "Copy Signature"
        else:
            title = "Edit Signature"

        self.setWindowTitle(title)
        self.setMinimumWidth(1000)
        self.setMinimumHeight(620)

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
        preview_btn = QPushButton("Preview")
        cancel_btn = QPushButton("Cancel")

        save_btn.clicked.connect(self.on_save)
        preview_btn.clicked.connect(self.on_preview)
        cancel_btn.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(preview_btn)
        button_layout.addWidget(cancel_btn)

        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

    def clear_global_id(self):
        """Очищает поле Global ID для режима копирования"""
        if 'global_id' in self.inputs:
            self.inputs['global_id'].clear()

    def clear_signature_name(self):
        """Очищает поле Signature Name для режима копирования"""
        if 'signature_name' in self.inputs:
            self.inputs['signature_name'].clear()

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

    def on_save(self):
        """Обработчик нажатия кнопки Save"""
        # Собираем данные из полей
        data = []

        # Основные поля
        basic_fields = ['global_id', 'signature_name', 'first_name', 'last_name',
                        'job', 'email', 'greet']
        for field in basic_fields:
            data.append(self.inputs[field].text().strip())

        # Контактные поля
        contact_fields = ['work_number', 'personal_number', 'social_number', 'cut_number']
        for field in contact_fields:
            data.append(self.inputs[field].text().strip())

        # Комбобоксы
        data.append(self.cb_hotel.currentIndex())
        data.append(self.cb_language.currentIndex())
        data.append(self.cb_type.currentIndex())

        # URL поля
        url_fields = ['banner_path', 'banner_url', 'site_url']
        for field in url_fields:
            data.append(self.inputs[field].text().strip())

        # Чекбоксы
        checkbox_fields = ['conf_greet', 'conf_fname', 'conf_job', 'conf_hotel',
                           'conf_phone_numbers', 'conf_mail', 'conf_banner',
                           'conf_site', 'conf_main_sig']
        for field in checkbox_fields:
            data.append(1 if self.checkboxes[field].isChecked() else 0)

        # Выполняем соответствующую операцию
        if self.mode == 'create' or self.mode == 'copy':
            success = self.db.insert_user(data)
            if success:
                # QMessageBox.information(self, "Success", "Signature created successfully!")
                self.accept()
            else:
                QMessageBox.warning(self, "Error", "Failed to create signature!")
        else:  # mode == 'edit'
            success = self.db.update_user(self.signature_id, data)
            if success:
                # QMessageBox.information(self, "Success", "Signature updated successfully!")
                self.accept()
            else:
                QMessageBox.warning(self, "Error", "Failed to update signature!")

    def on_preview(self):
        """Предварительный просмотр подписи"""
        # Собираем данные из полей (аналогично on_save, но без сохранения в БД)
        first_name = self.inputs['first_name'].text().strip()
        last_name = self.inputs['last_name'].text().strip()
        job = self.inputs['job'].text().strip()
        email = self.inputs['email'].text().strip()
        greet = self.inputs['greet'].text().strip()
        work_number = self.inputs['work_number'].text().strip()
        personal_number = self.inputs['personal_number'].text().strip()
        social_number = self.inputs['social_number'].text().strip()
        cut_number = self.inputs['cut_number'].text().strip()

        cb_hotel = self.cb_hotel.currentIndex()
        cb_language = self.cb_language.currentIndex()
        cb_type = self.cb_type.currentIndex()

        banner_path = self.inputs['banner_path'].text().strip()
        banner_url = self.inputs['banner_url'].text().strip()
        site_url = self.inputs['site_url'].text().strip()

        # Получаем значения чекбоксов
        conf_greet = 1 if self.checkboxes['conf_greet'].isChecked() else 0
        conf_fname = 1 if self.checkboxes['conf_fname'].isChecked() else 0
        conf_job = 1 if self.checkboxes['conf_job'].isChecked() else 0
        conf_hotel = 1 if self.checkboxes['conf_hotel'].isChecked() else 0
        conf_phone_numbers = 1 if self.checkboxes['conf_phone_numbers'].isChecked() else 0
        conf_mail = 1 if self.checkboxes['conf_mail'].isChecked() else 0
        conf_banner = 1 if self.checkboxes['conf_banner'].isChecked() else 0
        conf_site = 1 if self.checkboxes['conf_site'].isChecked() else 0
        conf_main_sig = 1 if self.checkboxes['conf_main_sig'].isChecked() else 0

        # Вызываем функцию создания подписи
        html_content = create_email_signature(
            first_name, last_name, job, email, greet,
            work_number, personal_number, social_number, cut_number,
            cb_hotel, cb_language, cb_type,
            banner_path, banner_url, site_url,
            conf_greet, conf_fname, conf_job, conf_hotel,
            conf_phone_numbers, conf_mail, conf_banner, conf_site
        )

        if html_content:
            # Способ 1: Открыть во встроенном браузере (QWebEngineView)
            try:
                from PyQt5.QtWebEngineWidgets import QWebEngineView

                preview_dialog = QDialog(self)
                preview_dialog.setWindowTitle("Signature Preview")
                preview_dialog.resize(800, 600)

                browser = QWebEngineView()
                browser.setHtml(html_content)

                layout = QVBoxLayout(preview_dialog)
                layout.addWidget(browser)

                preview_dialog.exec_()

            except ImportError:
                # Способ 2: Сохранить во временный файл и открыть в браузере
                import tempfile
                import webbrowser
                import os

                # Создаем временную директорию если её нет
                temp_dir = os.path.join(os.path.dirname(__file__), "temp_sig")
                os.makedirs(temp_dir, exist_ok=True)

                # Создаем временный файл
                temp_file = os.path.join(temp_dir, f"preview_{first_name}_{last_name}.htm")

                with open(temp_file, 'w', encoding='windows-1251') as f:
                    f.write(html_content)

                # Открываем в браузере по умолчанию
                webbrowser.open(f"file://{os.path.abspath(temp_file)}")

                # QMessageBox.information(self, "Preview",
                #                         "Signature opened in browser. Temporary file saved in temp_sig folder.")
        else:
            QMessageBox.warning(self, "Error", "Failed to generate signature preview")

if __name__ == '__main__':
    user_global_id = os.getlogin()
    config = configparser.ConfigParser()
    config.read('config.ini')
    list_users = config['admins']['admin']
    if user_global_id not in list_users:
        exit()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())