import os
import webbrowser


def create_email_signature(global_id,
                           first_name,
                           last_name,
                           job,
                           phone_numbers,
                           cut_phone,
                           email,
                           hotel,
                           banner,
                           banner_path,
                           link_href,
                           greet,
                           language,
                           type):

    # color and config from hotel
    hotel_config = None
    if language == 1:
        hotel_config = {
            1: {
                "color": "#441D61",
                "address": "620014, Россия, Екатеринбург, ул. Бориса Ельцина, 8",
                "hotel_name": ["Хаятт Ридженси Екатеринбург"]
            },
            2: {
                "color": "#ED7D31",
                "address": "620028, Россия, Екатеринбург, ул. Репина 1/2",
                "hotel_name": ["Хаятт Плейс Екатеринбург"]
            },
            3: {
                "color": "#441D61",
                "address": "620014, Россия, Екатеринбург, ул. Бориса Ельцина, 8",
                "hotel_name": ["Хаятт Ридженси Екатеринбург", "Хаятт Плейс Екатеринбург"]
            }
        }
    elif language == 2:
        hotel_config = {
            1: {
                "color": "#441D61",
                "address": "620014, Russia, Yekaterinburg, Borisa Yeltsina str. 8",
                "hotel_name": ["Hyatt Regency Yekaterinburg"]
            },
            2: {
                "color": "#ED7D31",
                "address": "620028, Russia, Yekaterinburg, Repina str. 1/2",
                "hotel_name": ["Hyatt Place Yekaterinburg"]
            },
            3: {
                "color": "#441D61",
                "address": "620014, Russia, Yekaterinburg, Borisa Yeltsina str. 8",
                "hotel_name": ["Hyatt Regency Yekaterinburg", "Hyatt Place Yekaterinburg"]
            }
        }
    if not hotel_config:
        return None
    config = hotel_config.get(hotel, None)
    
    # full name
    full_name = f'{first_name} {last_name}'.upper()

    # telephone numbers
    phones_html = ''
    work_phone = ''
    mobile_phone = ''
    whatsapp_phone = ''

    # types of numbers
    for phone_type, number in phone_numbers.items():
        if phone_type == 'T' and number:
            work_phone = number
        elif phone_type == 'M' and number:
            mobile_phone = number
        elif phone_type == 'WA' and number:
            whatsapp_phone = number

    # add phones
    
    # work phones
    if work_phone:
        phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'> T&nbsp;&nbsp;</span><span style='font-size:10.0pt;
font-family:"Arial",sans-serif;color:black'> {work_phone}&nbsp;&nbsp; <o:p></o:p></span></p>'''
    
    # mobile + wa
    if mobile_phone and whatsapp_phone:
        phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> М&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {mobile_phone}&nbsp;&nbsp;</span>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> WA&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {whatsapp_phone}&nbsp;&nbsp;</span>
<o:p></o:p>
</p>'''
    elif mobile_phone:
        phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'>M&nbsp;</span><span style='font-size:10.0pt;font-family:
"Arial",sans-serif;color:black'> {mobile_phone}&nbsp;&nbsp; <o:p></o:p></span></p>'''
    elif whatsapp_phone and not work_phone:
        phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'>WA</span><span style='font-size:10.0pt;font-family:
"Arial",sans-serif;color:black'> {whatsapp_phone}&nbsp;&nbsp; <o:p></o:p></span></p>'''
    else:
        phones_html = f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> T&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {work_phone}&nbsp;&nbsp;</span>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> WA&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {whatsapp_phone}&nbsp;&nbsp;</span>
<o:p></o:p>
</p>'''

    # greeting
    if greet:
        greeting = f'''
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Arial",sans-serif;
color:black'>{greet},<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Arial",sans-serif;
color:black'><o:p>&nbsp;</o:p></span></p>
'''
    else:
        greeting = ''

    # f name
    full_name_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:10.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:{config["color"]};text-transform:
uppercase'>{full_name}<o:p></o:p></span></b></p>
'''

    # job
    job_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><span style='font-size:10.0pt;mso-bidi-font-size:
9.0pt;line-height:120%;font-family:"Arial",sans-serif;mso-bidi-font-weight:
bold'>{job.capitalize()}<o:p></o:p></span></p>
'''

    # space_after_job
    space_after_job = '''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='mso-ascii-font-family:Calibri;mso-hansi-font-family:Calibri;mso-bidi-font-family:
Calibri'><o:p>&nbsp;</o:p></span></p>
'''

    # hotel and address
    hotel_and_address = ''
    for item in config["hotel_name"]:
        hotel_and_address += f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:"Arial",sans-serif;
color:{config["color"]}'>{item.replace("<br>", "<o:p></o:p></span></p><p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:\"Arial\",sans-serif; color:{config[\"color\"]}'>")}<o:p></o:p></span></p>
'''
    hotel_and_address += f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:"Arial",sans-serif'>
{config["address"]}<o:p></o:p></span></p>
'''

    # space_after_address
    space_after_address = '''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>
'''

    # email
    email_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}'> E&nbsp;&nbsp;&nbsp;</span><span
lang=EN-US style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-ansi-language:
EN-US'><a href="mailto:{email}">{email}</a></span><span
style='font-size:10.0pt;font-family:"Arial",sans-serif'><o:p></o:p></span></p>
'''
    # banner
    banner_html = ''
    if banner == 1:
        try:
            with open(banner_path, 'rb'):
                banner_html = f'''
                <p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
                line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
                line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>
    
                <p class=MsoNormal><a href="{link_href}">
                <img border=0 width=779 height=136 src="{banner_path}" style="border:none;">
                </a><o:p></o:p></p>'''
        except FileNotFoundError:
            banner_html = ''

    # full html
    if type != 1:
        space_after_job = ''
        hotel_and_address =  ''
        space_after_address = ''
        phones_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'></span><span style='font-size:10.0pt;
font-family:"Arial",sans-serif;color:black'>{cut_phone}&nbsp;&nbsp; <o:p></o:p></span></p>'''
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

<p class=MsoNormal><o:p>&nbsp;</o:p></p>
</div>
</body>
</html>'''

    return html_signature


def save_signature_to_file(html_content, global_id, user_global_id):
    signatures_path = rf'C:\Users\{user_global_id}\AppData\Roaming\Microsoft\Signatures'
    filename = f'{global_id}.htm'
    full_path = os.path.join(signatures_path, filename)
    with open(full_path, 'w', encoding='windows-1251') as f:
        f.write(html_content)
    print(f'Подпись сохранена по пути: {full_path}')
    webbrowser.open(os.path.abspath(full_path))


if __name__ == "__main__":

    user_global_id = os.environ['USERNAME']

    first_phones = {
        'T': '+7 343 123 1234',
        'M': '+7 963 123 1234',
        'WA': ''
    }

    # global_id, first_name, last_name, job, phone_numbers, cut_phone,
    # email, hotel, banner, banner_path, link_href, greet, language, type

    # global id
    # имя
    # фамилия
    # должность
    # почта
    # приветствие

    # рабочий номер
    # личный номер
    # номер whatsapp
    # короткий номер

    # приветствие

    # отель
    # язык
    # тип

    # путь до баннера
    # ссылка для баннера
    # сайт

    # чекбоксы #

    # убрать баннер
    # убрать имя
    # убрать должность
    # убрать отели
    # убрать номера телефонов
    # убрать адрес почты

    # добавить сайт
    # добавить приветствие

    first = [
        '1234567',                                 # global_id
        'Иван',                                    # first_name
        'Иванов',                                # last_name
        'Специалист по информационным системам',   # job
        first_phones,                              # phone_numbers
        '+7 963 123 1234 (*1234 / 61234)',         # cut_phone
        'ivan.ivanov@example.ru',   # email
        3,                                         # hotel
        1,                                         # banner
        r'D:\scripts\py\actual\auto_update_signatures_outlook\banner.jpg',    # banner_path
        r'https://ya.ru',                                                 # link_href
        'С уважением',                             # greet
        1,                                         # language
        1                                          # type
    ]

    second_phones = {
        'T': '',
        'M': '',
        'WA': '+7 963 123 1234'
    }

    # global_id, first_name, last_name, job, phone_numbers, cut_phone,
    # email, hotel, banner, banner_path, link_href, greet, language, type

    second = [
        '1234568',  # global_id
        'иван',  # first_name
        'иВанов',  # last_name
        'специалист по информационным системам',  # job
        second_phones,  # phone_numbers
        '+7 963 123 1234 (*1234 / 61234)', # cut_phone
        'ivan.ivanov@example.ru',  # email
        2,  # hotel
        1,  # banner
        r'D:\scripts\py\actual\auto_update_signatures_outlook\banner.png',
        # banner_path
        r'https://ya.ru',  # link_href
        'С уважением',  # greet
        2,  # language
        2  # type
    ]

    users = [first, second]

    for user in users:

        global_id, first_name, last_name, job, phone_numbers, cut_phone, email, hotel, banner, banner_path, link_href, greet, language, type = user # first / second

        values = [
            global_id,
            first_name,
            last_name,
            job,
            phone_numbers,
            cut_phone,
            email,
            hotel,
            banner,
            banner_path,
            link_href,
            greet,
            language,
            type
        ]

        for i in values:
            print(i)

        # type (1 - full, 2 - cut)
        # hotel (1 - Hyatt Regency , 2 - Hyatt Place, 3 - both)
        # banner (1 - ON, 2 - OFF)
        # language (1 - ru, 2 - en)

        signature_html = create_email_signature(
            global_id=global_id,
            first_name=first_name,
            last_name=last_name,
            job=job,
            phone_numbers=phone_numbers,
            cut_phone=cut_phone,
            email=email,
            hotel=hotel,
            banner=banner,
            banner_path=banner_path,
            link_href=link_href,
            greet=greet,
            language=language,
            type=type
        )

        if signature_html:
            save_signature_to_file(signature_html, global_id, user_global_id)