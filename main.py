import os
import webbrowser
import win32security
import winreg

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

    # color and config from hotel
    hotel_config = None
    if cb_language == 1:
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
    elif cb_language == 2:
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
    config = hotel_config.get(cb_hotel, None)
    
    # full name
    full_name = f'{first_name} {last_name}'.upper()

    # telephone numbers
    try:
        with open('tags.txt', 'r', encoding='utf-8') as f:
            tags = []
            for line in f:
                if not line.startswith('#'):
                    tags.append(line.rstrip('\n'))
                    if len(tags) == 3:
                        break
        work_tag, pers_tag, soc_tag = tags[0], tags[1], tags[2] if len(tags) == 3 else ['', '', '']
    except Exception:
        work_tag, pers_tag, soc_tag = 'T', 'M', 'WA'


    phones_html = ''

    # add phones
    if conf_phone_numbers == 1:
    # work phones
        if work_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'> {work_tag}&nbsp;&nbsp;</span><span style='font-size:10.0pt;
font-family:"Arial",sans-serif;color:black'> {work_number}&nbsp;&nbsp; <o:p></o:p></span></p>'''
    
        # mobile + wa
        if personal_number and social_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> {pers_tag}&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {personal_number}&nbsp;&nbsp;</span>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> {soc_tag}&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {social_number}&nbsp;&nbsp;</span>
<o:p></o:p>
</p>'''
        elif personal_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'>{pers_tag}&nbsp;</span><span style='font-size:10.0pt;font-family:
"Arial",sans-serif;color:black'> {personal_number}&nbsp;&nbsp; <o:p></o:p></span></p>'''
        elif social_number and not work_number:
            phones_html += f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'>{soc_tag}</span><span style='font-size:10.0pt;font-family:
"Arial",sans-serif;color:black'> {social_number}&nbsp;&nbsp; <o:p></o:p></span></p>'''
        else:
            phones_html = f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> {work_tag}&nbsp;</span>
<span style='font-size:10.0pt; font-family:"Arial",sans-serif;color:black'> {work_number}&nbsp;&nbsp;</span>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}; mso-bidi-font-weight:bold'> {soc_tag}&nbsp;</span>
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
line-height:120%;font-family:"Arial",sans-serif;color:{config["color"]};text-transform:
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
        for item in config["hotel_name"]:
            hotel_and_address += f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:"Arial",sans-serif;
color:{config["color"]}'>{item.replace("<br>", "<o:p></o:p></span></p><p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:\"Arial\",sans-serif; color:{config[\"color\"]}'>")}<o:p></o:p></span></p>'''
        hotel_and_address += f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;mso-bidi-font-size:11.0pt;font-family:"Arial",sans-serif'>
{config["address"]}<o:p></o:p></span></p>'''

        # space_after_address
        space_after_address = '''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>'''

    # email
    email_html = ''
    if conf_mail == 1:
        email_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'>
<span style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]}'> E&nbsp;&nbsp;&nbsp;</span><span
lang=EN-US style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-ansi-language:
EN-US'><a href="mailto:{email}">{email}</a></span><span
style='font-size:10.0pt;font-family:"Arial",sans-serif'><o:p></o:p></span></p>'''

    # banner
    banner_html = ''
    if conf_banner == 1:
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

                with open(banner_path, 'rb'):
                    banner_html = f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>

<p class=MsoNormal><a href="{banner_url}">
<img border=0 width=779 height=136 src="{banner_path}" style="border:none;">
</a><o:p></o:p></p>'''

    # site
    site_html = ''
    if conf_site == 1:
        site_html = f'''
<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph;
line-height:120%;text-autospace:none'><b><span style='font-size:9.0pt;
line-height:120%;font-family:"Arial",sans-serif;color:#151F6D'><o:p>&nbsp;</o:p></span></b></p>
        
<p class='MsoNormal' style='text-align:justify; text-justify:inter-ideograph; font-size:10.0pt; font-family:Arial, sans-serif; color:{config["color"]};'>
    <a href='{site_url}' style='color: inherit; text-decoration: none;'>{site_url}</a>
</p>'''

    # full html
    if cb_type != 1:
        space_after_job = ''
        hotel_and_address =  ''
        space_after_address = ''
        phones_html = f'''<p class=MsoNormal style='text-align:justify;text-justify:inter-ideograph'><span
style='font-size:10.0pt;font-family:"Arial",sans-serif;color:{config["color"]};
mso-bidi-font-weight:bold'></span><span style='font-size:10.0pt;
font-family:"Arial",sans-serif;color:black'>{cut_number}&nbsp;&nbsp; <o:p></o:p></span></p>'''
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
    signatures_path = rf'C:\Users\{user_global_id}\AppData\Roaming\Microsoft\Signatures'
    if not os.path.exists(signatures_path):
        os.makedirs(signatures_path)
    filename = f'{global_id}-{signature_name}.htm'
    full_path = os.path.join(signatures_path, filename)
    with open(full_path, 'w', encoding='windows-1251') as f:
        f.write(html_content)
    print(full_path)
    webbrowser.open(os.path.abspath(full_path))

def set_outlook_signature(sid, signature_name, global_id):
    reg_path = rf"{sid}\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002"
    signature_name = f'{global_id}-{signature_name}'
    try:
        key = winreg.OpenKey(winreg.HKEY_USERS, reg_path, 0, winreg.KEY_WRITE)

        winreg.SetValueEx(key, "New Signature", 0, winreg.REG_SZ, signature_name)
        winreg.SetValueEx(key, "Reply-Forward Signature", 0, winreg.REG_SZ, signature_name)

        winreg.CloseKey(key)
        return True

    except Exception:
        pass


if __name__ == "__main__":

    user_global_id = os.getlogin() # os.environ['USERNAME']

    user_info = win32security.LookupAccountName(None, os.getlogin())
    sid = win32security.ConvertSidToStringSid(user_info[0])

    print(sid)

    first = [
        'base',                                     # signature_name
        '1234567',                                  # global_id
        'Иван',                                     # first_name
        'Иванов',                                   # last_name
        'Специалист по информационным системам',    # job
        'ivan.ivanov@example.ru',                   # email
        'С уважением',                              # greet
        '+7 343 123 1234',                          # work_number
        '+7 963 123 1234',                          # personal_number
        '',                                         # social_number
        '+7 963 123 1234 (*1234 / 61234)',          # cut_number
        3,                                          # cb_hotel (1 - Hyatt Regency , 2 - Hyatt Place, 3 - both)
        1,                                          # cb_language (1 - ru, 2 - en)
        2,                                          # cb_type (1 - full, 2 - cut)
        r'D:\scripts\py\actual\auto_update_signatures_outlook\banner.jpg',  # banner_path
        r'https://ya.ru',                           # banner_url
        r'https://ya.ru',                           # site_url
        1,                                          # conf_greet (1 - enable, 2 - disable)
        1,                                          # conf_fname (1 - enable, 2 - disable)
        1,                                          # conf_job (1 - enable, 2 - disable)
        1,                                          # conf_hotel (1 - enable, 2 - disable)
        1,                                          # conf_phone_numbers (1 - enable, 2 - disable)
        1,                                          # conf_mail (1 - enable, 2 - disable)
        1,                                          # conf_banner (1 - enable, 2 - disable)
        1,                                          # conf_site (1 - enable, 2 - disable)
        1,                                          # conf_main_sig (1 - enable, 2 - disable)
    ]

    users = [first, ]

    for user in users:

        (signature_name, global_id, first_name, last_name, job, email, greet,
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

        values = [
            signature_name, global_id, first_name, last_name, job, email, greet,
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
        ]

        for i in values:
            ... # print(i)

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