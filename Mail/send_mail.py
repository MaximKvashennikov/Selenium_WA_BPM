import win32com.client as win32
from InputData.robot_dict import region_dict
import re


class SendMail:
    def __init__(self, win32, text_wa, region, start_work, start_time, rrl_list_sw_file):
        self.win32 = win32
        self.text_wa = text_wa
        self.region = region
        self.start_work = start_work
        self.start_time = start_time
        self.rrl_list_sw_file = rrl_list_sw_file

    def send_mail(self):
        outlook = self.win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail_to_reg = region_dict[self.region]['mail'].replace(",", ";")
        print(mail_to_reg)
        mail.To = mail_to_reg
        # mail.To = "denis.kozhin@tele2.ru; Nikolay.Pogodin@tele2.ru"
        # mail.CC = "Nikolay.Pogodin@tele2.ru"
        mail.Subject = 'Согласование плановых работ {text_wa}'.format(
            text_wa=self.text_wa,
        )

        self.html_body(mail)
        mail.Send()

    def html_body(self, mail):
        body_mail = '''
        <html>
        <head>
            <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
            <title>HTML Template</title>
            <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
            <style type="text/css">
            body {
              width: 100% !important;
              -webkit-text-size-adjust: 100%;
              -ms-text-size-adjust: 100%;
              margin: 0;
              padding: 0;
              line-height: 100%;
            }
        
            [style*="Open Sans"] {
              font-family: 'Open Sans', arial, sans-serif !important;
            }
        
            [style*="Arial"] {
              font-family: 'Arial', Open Sans, sans-serif !important;
            }
        
            img {
              outline: none;
              text-decoration: none;
              border: none;
              -ms-interpolation-mode: bicubic;
              max-width: 100% !important;
              margin: 0;
              padding: 0;
              display: block;
            }
        
            table td {
              border-collapse: collapse;
            }
        
            table {
              border-collapse: collapse;
              mso-table-lspace: 0pt;
              mso-table-rspace: 0pt;
            }
        
            .art-img {
              width: 520px;
              height: 180px;
            }
        
            .img-280 {
              display: none;
            }
        
            .logo {
              width: 187px;
              padding-top: 30px;
              padding-bottom: 28px;
            }
        
            .white {
              background-color: #ffffff !important;
            }
        
            @media(max-width:620px) {
        
              .table-600 {
                width: 280px !important;
              }
        
              .table-520 {
                width: 244px !important;
              }
        
              .logo {
                width: 114px;
                padding-top: 17px;
                padding-bottom: 17px;
              }
        
              .hero-img {
                width: 280px !important;
                height: 300px !important;
              }
        
              .art-img {
                width: 244px !important;
                height: 84px !important;
              }
        
              .img-600 {
                display: none;
              }
        
              .img-280 {
                display: block;
              }
        
              .icon-table {
                display: inline;
              }
            }
        
        
            </style>
        </head>
        
        <body style="margin: 0; padding: 0;">
        <table cellpadding="0" cellspacing="0" width="100%" bgcolor="#ededed">
        
            <tr>
                <td>
                    <table align="center" cellpadding="0" cellspacing="0" width="600" class="table-600" bgcolor="#000000">
                        <tr>
        
                        </tr>
                    </table>
                </td>
            </tr>
        
            <tr>
                <td align="center">
                    <table width="600" cellpadding="0" cellspacing="0" class="table-600 white" bgcolor="#ffffff">
                        <tr>
                            <td align="center">
                                <table width="520" class="table-520" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td>
                                            <h1
                                                    style="font-family: Arial, Open Sans, sans-serif;color:#515151;margin-top: 32px;margin-bottom: 13px; font-size: 20px;line-height: 24px;">
                                                Изменение настроек/конфигурации РРЛ
                                            </h1>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <p style="font-family: Arial, Open Sans, sans-serif;color:#515151;margin-top:13px;margin-bottom:13px;font-size: 16px;line-height: 24px;">
                                                Добрый день.
                                                <br>
                                                <br>
                                                В регионе <% region %> будут проводиться работы по изменению
                                                настроек/конфигурации РРЛ под номером <% text_wa %>.
                                                <br>
                                                Начало работ: <% start_work %> <% start_time %>
                                                <br>
                                                Список затронутых РРЛ:
                                                <br>
                                                <% rrl_list_sw_file %>
                                                <br>
                                                <b>
                                                    Уведомляем вас о необходимости их согласования.
                                                </b>
                                                <br>
                                                В случае отклонения просьба сообщить Transport.CP_Access_Exploitation.
                                            </p>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        
            <tr>
                <td align="center">
                    <table width="600" cellpadding="0" cellspacing="0" class="table-600" bgcolor="#000000">
                        <tr>
                            <td align="center">
                                <table width="520" class="table-520" cellpadding="0" cellspacing="0">
        
                                </table>
                            </td>
                        </tr>
        
                    </table>
                </td>
            </tr>
        
        
        </table>
        
        </body>
        
        </html>
                '''
        body_mail = re.sub(r'<% text_wa %>', self.text_wa, body_mail)
        body_mail = re.sub(r'<% region %>', self.region, body_mail)
        body_mail = re.sub(r'<% start_work %>', self.start_work, body_mail)
        body_mail = re.sub(r'<% start_time %>', self.start_time, body_mail)
        body_mail = re.sub(r'<% rrl_list_sw_file %>', self.rrl_list_sw_file, body_mail)

        mail.HTMLBody = body_mail


if __name__ == "__main__":
    SendMail(win32=win32, text_wa="WA...", region="Центр - Калуга", start_work="21.09.2021", start_time="0:00").send_mail()
