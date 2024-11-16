import imaplib
import ssl

def verify_hotmail(email, password):
    host = "outlook.office365.com"
    port = 993  # IMAP SSL port
    
    try:

        context = ssl.create_default_context()
        
        imap = imaplib.IMAP4_SSL(host, port, ssl_context=context)
        
        imap.login(email, password)
        return True, "Login successful"
        
    except imaplib.IMAP4.error as e:
        return False, "Login failed: {}".format(str(e))
    finally:
        try:
            imap.logout()
        except:
            pass
