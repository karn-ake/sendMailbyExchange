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
            body='Python test script',
            to_recipients=[
                exchangelib.Mailbox(email_address='karn-ake.r@n2nconnect.com'),
            ],
            #cc_recipients=['thdc@n2nconnect.com'],
        )
        with open('User Report _2021_06.zip','rb') as attachFile:
            binary = attachFile.read()
        attachments = exchangelib.FileAttachment(name='User Report _2021_06.zip', content=binary)
        messages.attach(attachments)
        messages.send()

    connectedExchange()