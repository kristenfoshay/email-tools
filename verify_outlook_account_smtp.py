import smtplib
import ssl

def verify_outlook_account(email, password):

    host = "smtp.office365.com"
    port = 587
    
    try:
      
        server = smtplib.SMTP(host, port)
        server.starttls(context=ssl.create_default_context())
        
        server.login(email, password)
        return True, "Login successful"
        
    except smtplib.SMTPAuthenticationError:
        return False, "Invalid credentials"
    except Exception as e:
        return False, "An error occurred: {}".format(str(e))
    finally:
        try:
            server.quit()
        except:
            pass

# Example usage:
if __name__ == "__main__":
    email = input("Enter your email: ")
    password = input("Enter your password: ")
    
    success, message = verify_outlook_account(email, password)
    print(message)
