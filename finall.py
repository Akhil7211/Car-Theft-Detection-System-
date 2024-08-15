# Importing the modules
import imaplib  # For connecting to the email server
import email  # For parsing the email messages
import pyttsx3  # For converting text to speech
import time  # For adding a delay between each check

# Creating an IMAP object and logging in with your credentials
imap = imaplib.IMAP4_SSL("imap.gmail.com")
imap.login("ecsproject2023@gmail.com", "fthwnahozzuooxmc")

# Creating a text-to-speech engine and setting the voice properties
engine = pyttsx3.init()
engine.setProperty("rate", 150)  # Adjusting the speed of speech
engine.setProperty("volume", 0.8)  # Adjusting the volume of speech


# Defining a function to read an email message as voice
def read_email(message):
    # Extracting the subject and the body of the message
    subject = message["Subject"]
    body = ""
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                body += part.get_payload(decode=True).decode()
    else:
        content_type = message.get_content_type()
        if content_type == "text/plain":
            body += message.get_payload(decode=True).decode()

    # Speaking out the subject and the body of the message
    engine.say(f"You have a new email from {message['From']}. The subject is {subject}. The body is {body}")
    engine.runAndWait()


# Defining a loop to check for new messages every 10 seconds
while True:
    # Selecting the inbox folder and searching for unread messages
    imap.select("INBOX")
    status, messages = imap.search(None, "UNSEEN")

    # Getting all unread messages and reading them as voice
    messages = messages[0].split()
    for msg in messages:
        status, data = imap.fetch(msg, "(RFC822)")
        message = email.message_from_bytes(data[0][1])
        read_email(message)

        # Adding a delay before checking again
        time.sleep(10)
