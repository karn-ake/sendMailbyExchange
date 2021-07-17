import exchangelib

USERNAME = 'n2n\\nawapong'
PASSWORD = 'N2Nconnect123'
EXCHANGE_SERVER = 'mail.n2nconnect.com'
SENDER = 'nawapong.a@n2nconnect.com'
TO_RECIPIENTS = 'karn-ake.r@n2nconnect.com'
CC_RECIPIENTS = 'thdc@n2nconnect.com'
SUBJECT = 'Test send mail to and cc from script'
BODY = """<html>
            <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
            <style type="text/css" style="font-family:Tahoma;font-size:15pt;"></style>
            </head>
            <body>
            <b>Dear All,</b>
            
            <p>Here are test sending messages.</p>
            
            <i>Best Regards,</i>
            <p>Thailand Operation Team</p></body></html>
            """
ATTACHMENT = 'Trade Value 202106.xlsx'

class sendEmail:

    def __init__(self, username, password, exchange_server, sender, to_recipients, cc_recipients, subject, body, attachment):
        self.username = username
        self.password = password
        self.exchange_server = exchange_server
        self.sender = sender
        self.to_recipients = to_recipients
        self.cc_recipients = cc_recipients
        self.subject = subject
        self.body = body
        self.attachment = attachment

    def connectedExchange(self):
        credentials = exchangelib.Credentials(self.username, self.password)
        config = exchangelib.Configuration(server=self.exchange_server, credentials=credentials)
        account = exchangelib.Account(primary_smtp_address=self.sender, autodiscover=False, config=config, access_type=exchangelib.DELEGATE)
        messages = exchangelib.Message(
            account=account,
            subject=self.subject,
            body=exchangelib.HTMLBody(self.body),
            to_recipients=[
                exchangelib.Mailbox(email_address=self.to_recipients),
            ],
            cc_recipients=[self.cc_recipients],
        )
        with open(self.attachment,'rb') as attachFile:
            binary = attachFile.read()
        attachments = exchangelib.FileAttachment(name=self.attachment, content=binary)
        messages.attach(attachments)
        messages.send()

daily_report = sendEmail(USERNAME, PASSWORD, EXCHANGE_SERVER, SENDER, TO_RECIPIENTS, CC_RECIPIENTS, SUBJECT, BODY, ATTACHMENT)
daily_report.connectedExchange()