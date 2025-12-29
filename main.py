import os
import webbrowser
import win32security
import winreg
import pyodbc
import configparser
import sys
import threading
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu, QAction)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QObject, pyqtSignal
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QPixmap

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
    config.read('config.ini', encoding='utf-8')
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
            # Безопасное получение расширения файла
            if banner_path.lower().endswith('.png') or banner_path.lower().endswith('.jpg'):
                try:
                    with open(banner_path, 'rb'):
# old
#                         banner_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
# line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
# line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>
#
# <p class=MsoNormal><a href="{banner_url}">
# <img border=0 width=779 height=136 src="{banner_path}" style="border:none;">
# </a><o:p></o:p></p>'''
                        banner_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>

<p class=MsoNormal>
<a href="{banner_url}">
 <img border=0 width=779 height=136 src="{banner_path}" style="border:none; display:block;">
</a>
<o:p></o:p>
</p>'''
                except:
                    base, ext = os.path.splitext(banner_path)

                    if ext.lower() == '.png':
                        new_banner_path = base + '.jpg'
                    else:
                        new_banner_path = base + '.png'
                    try:
                        with open(new_banner_path, 'rb'):
                            banner_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>

<p class=MsoNormal>
<a href="{banner_url}">
 <img border=0 width=779 height=136 src="{new_banner_path}" style="border:none; display:block;">
</a>
<o:p></o:p>
</p>'''
                    except Exception as e:
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
        hotel_and_address =  ''
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

def save_signature_to_file(html_content, signature_name, global_id, user_global_id):
    # read ini file
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    signatures_path = rf'C:\Users\{user_global_id}' + config['settings']['signatures_path']
    if not os.path.exists(signatures_path):
        os.makedirs(signatures_path)
    filename = f'{global_id}-{signature_name}.htm'
    full_path = os.path.join(signatures_path, filename)
    with open(full_path, 'w', encoding='windows-1251') as f:
        f.write(html_content)
    # print(full_path)
    # webbrowser.open(os.path.abspath(full_path))

def set_outlook_signature(sid, signature_name, global_id):
    # read ini file
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    reg_path = rf'{sid}' + config['settings']['reg_path']
    signature_name = f'{global_id}-{signature_name}'
    try:
        key = winreg.OpenKey(winreg.HKEY_USERS, reg_path, 0, winreg.KEY_WRITE)

        winreg.SetValueEx(key, 'New Signature', 0, winreg.REG_SZ, signature_name)
        winreg.SetValueEx(key, 'Reply-Forward Signature', 0, winreg.REG_SZ, signature_name)

        winreg.CloseKey(key)
        return True

    except Exception:
        pass

class DatabaseManager:
    def __init__(self):
        # read ini file
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')
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
        query = 'SELECT * FROM signatures WHERE global_id = ?'
        return self.execute_query(query, [user_id], fetch_all=True)


class TrayApp(QObject):
    update_signal = pyqtSignal()
    update_complete_signal = pyqtSignal(list)  # Новый сигнал для завершения обновления
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    conf_notification_frequency = config['settings']['frequency']

    def __init__(self):
        super().__init__()
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)

        # Таймер автообновления
        self.timer = QTimer()
        self.timer.timeout.connect(self.auto_update)

        # Создаем иконку в трее
        self.tray_icon = QSystemTrayIcon()

        # Здесь нужна иконка - создай icon.ico или укажи путь к своей
        try:
            self.tray_icon.setIcon(QIcon("icon.ico"))
        except:
            # Если иконки нет, создадим пустую
            self.tray_icon.setIcon(QIcon())

        self.tray_icon.setToolTip("Outlook Signature Updater")

        # Создаем меню
        self.menu = QMenu()

        # Пункт "Обновить"
        update_action = QAction("Обновить", self.menu)
        update_action.triggered.connect(self.update_signatures)
        self.menu.addAction(update_action)

        # Пункт "Выйти"
        exit_action = QAction("Выйти", self.menu)
        exit_action.triggered.connect(self.exit_app)
        self.menu.addAction(exit_action)

        self.tray_icon.setContextMenu(self.menu)
        self.tray_icon.show()

        # Соединяем сигналы
        self.update_signal.connect(self.run_main)
        self.update_complete_signal.connect(self.show_update_notification)

    def auto_update(self):
        self.update_signatures()

    def update_signatures(self):
        # Запускаем в отдельном потоке, чтобы не блокировать GUI
        thread = threading.Thread(target=self.run_main)  # Изменено здесь
        thread.daemon = True
        thread.start()

    def run_main(self):
        # Запускаем оригинальную функцию main и передаем результат
        updated_signatures = main_with_return()
        # Отправляем сигнал с результатами
        self.update_complete_signal.emit(updated_signatures)

    def show_update_notification(self, signatures):

        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')
        conf_notification = config['settings']['notification']

        if signatures:
            signature_names = "\n".join(signatures)
            message = f"Обновлены подписи:\n{signature_names}"
        else:
            message = "Подписи не были обновлены"

        if conf_notification.strip()  == '1' and message != "Подписи не были обновлены":
            # Показываем уведомление на 5 секунд
            self.tray_icon.showMessage(
                "Outlook Signature Updater",
                message,
                QSystemTrayIcon.NoIcon,
                5000  # 5 секунд
            )

    def exit_app(self):
        if self.timer.isActive():
            self.timer.stop()
        self.tray_icon.hide()
        self.app.quit()
        sys.exit(0)

    def run(self):
        # Запускаем приложение сразу при старте
        self.update_signatures()
        self.timer.start(int(self.conf_notification_frequency) * 60 * 1000)  # 60 mins
        sys.exit(self.app.exec_())


def main_with_return():
    user_global_id = os.getlogin()
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    list_users = config['users']['list_users']
    black_list = config['users']['black_list']
    if user_global_id not in list_users and list_users or user_global_id in black_list and black_list:
        exit()
    updated_signatures = []  # Список для хранения имен обновленных подписей

    db = DatabaseManager()
    all_users_data = db.get_user_data(user_global_id)

    users = []
    if all_users_data:
        for row in all_users_data:
            users.append(row)

    # print(db.get_user_data(user_global_id))

    # sid
    user_info = win32security.LookupAccountName(None, os.getlogin())
    sid = win32security.ConvertSidToStringSid(user_info[0])

    for user in users:
        (signature_id, global_id, signature_name, first_name, last_name, job, email, greet,
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
         conf_main_sig
         ) = user

        signature_html = create_email_signature(
            first_name=first_name,
            last_name=last_name,
            job=job,
            email=email,
            greet=greet,
            work_number=work_number,
            personal_number=personal_number,
            social_number=social_number,
            cut_number=cut_number,
            cb_hotel=cb_hotel,
            cb_language=cb_language,
            cb_type=cb_type,
            banner_path=banner_path,
            banner_url=banner_url,
            site_url=site_url,
            conf_greet=conf_greet,
            conf_fname=conf_fname,
            conf_job=conf_job,
            conf_hotel=conf_hotel,
            conf_phone_numbers=conf_phone_numbers,
            conf_mail=conf_mail,
            conf_banner=conf_banner,
            conf_site=conf_site,
        )

        if signature_html:
            save_signature_to_file(signature_html, signature_name, global_id, user_global_id)
            if conf_main_sig == 1:
                set_outlook_signature(sid, signature_name, global_id)
            updated_signatures.append(f'{user_global_id}-{signature_name}')

    return updated_signatures

def main():
    main_with_return()


if __name__ == "__main__":
    tray_app = TrayApp()
    tray_app.run()