import imaplib
import email
import ssl
from email.header import decode_header

def read_email_messages(email_address, password, limit=5):

    host = "outlook.office365.com"
    port = 993
    
    try:
        
        context = ssl.create_default_context()
        imap = imaplib.IMAP4_SSL(host, port, ssl_context=context)
        imap.login(email_address, password)
        print("Login successful\n")
        imap.select("INBOX")
        _, message_numbers = imap.search(None, "ALL")
        message_numbers = message_numbers[0].split()
        
        for num in reversed(message_numbers[:limit]):
            try:
                _, msg_data = imap.fetch(num, "(RFC822)")
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)
                
                subject = decode_header(email_message["Subject"])[0][0]
                if isinstance(subject, bytes):
                    subject = subject.decode()
                
                from_ = decode_header(email_message.get("From", ""))[0][0]
                if isinstance(from_, bytes):
                    from_ = from_.decode()
                
                date = email_message["Date"]
                
                print("=" * 50)
                print(f"From: {from_}")
                print(f"Subject: {subject}")
                print(f"Date: {date}")
                
                if email_message.is_multipart():
                    # Handle multipart messages
                    for part in email_message.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode()
                            print("\nBody:")
                            print(body[:300] + "..." if len(body) > 300 else body)
                            break
                else:
                    body = email_message.get_payload(decode=True).decode()
                    print("\nBody:")
                    print(body[:300] + "..." if len(body) > 300 else body)
                
                print("\n")
                
            except Exception as e:
                print(f"Error reading message: {str(e)}")
                continue
        
    except imaplib.IMAP4.error as e:
        print(f"Login failed: {str(e)}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        try:
            imap.logout()
        except:
            pass

if __name__ == "__main__":
    email_address = input("Enter your email: ")
    password = input("Enter your password: ")
    num_messages = input("How many recent messages to show? (default 5): ")
    
    try:
        num_messages = int(num_messages)
    except:
        num_messages = 5
    
    read_email_messages(email_address, password, num_messages)
