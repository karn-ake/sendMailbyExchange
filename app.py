import exchangelib

class sendEmail:
    def __init__(self):
        pass

    def connectedExchange():
        credentials = exchangelib.Credentials('n2n\\nawapong', 'N2Nconnect123')
        config = exchangelib.Configuration(server='mail.n2nconnect.com', credentials=credentials)
        account = exchangelib.Account(primary_smtp_address='nawapong.a@n2nconnect.com', autodiscover=False, config=config, access_type=exchangelib.DELEGATE)
        messages = exchangelib.Message(
            account=account,
            subject='Test send mail to and cc from script',
            body=exchangelib.HTMLBody("""<html>
            <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
            <style type="text/css" style="font-family:Tahoma;font-size:15pt;"></style>
            </head>
            <body>
            <b>Dear All,</b>
            
            <p>Here are test sending messages.</p>
            
            <i>Best Regards,</i>
            <p>Thailand Operation Team</p></body></html>
            """),
            to_recipients=[
                exchangelib.Mailbox(email_address='karn-ake.r@n2nconnect.com'),
            ],
            #cc_recipients=['thdc@n2nconnect.com'],
        )
        with open('Trade Value 202106.xlsx','rb') as attachFile:
            binary = attachFile.read()
        attachments = exchangelib.FileAttachment(name='Trade Value 202106.xlsx', content=binary)
        messages.attach(attachments)
        messages.send()

    connectedExchange()