import win32com.client as win32

SENDER_EMAIL = "eerxuan@outlook.com"
RECEIVER_EMAIL = "eerxuan@gmail.com"


def send_email(receiver_email, message):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)


    mail.Subject = message
    mail.To = receiver_email

    mail.Body = message

    mail.Send()


def main():
    send_email(RECEIVER_EMAIL, "Hello gmail !")


if __name__ == "__main__":
    main()