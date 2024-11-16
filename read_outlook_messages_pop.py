import poplib
from email.parser import Parser

HOST = 'pop-mail.outlook.com'
PORT = 995

USERNAME = 'email@outlook.com'
PASSWORD = 'password'

def read_outlook_messages():
    try:
        server = poplib.POP3_SSL(HOST, PORT)
        server.user(USERNAME)
        server.pass_(PASSWORD)

        num_messages = len(server.list()[1])
        print(f'Number of messages: {num_messages}\n')

        for i in range(num_messages):
            response, lines, octets = server.retr(i + 1)
            message_content = b'\r\n'.join(lines).decode('utf-8')
            message = Parser().parsestr(message_content)
            print(f'From: {message["from"]}')
            print(f'Subject: {message["subject"]}')
            print('-' * 60)

        server.quit()
   
    except Exception as e:
        print(f'An error occurred: {e}')

if __name__ == "__main__":
    read_hotmail_messages()
