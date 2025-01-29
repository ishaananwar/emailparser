import imaplib
import email
from email.header import decode_header
import requests
from typing import Dict, Any
import os
from flask import Flask, jsonify
from contextlib import contextmanager

class EmailTicketParser:
    def __init__(self, imap_host: str, username: str, password: str, webhook_url: str):
        self.imap_host = imap_host
        self.username = username
        self.password = password
        self.webhook_url = webhook_url
        self.blocked_senders = [
            "noreply@ingeniumstem.org",
            "Mailer-Daemon@mx1.mxfilter.net"
        ]

    @contextmanager
    def connect(self):
        """Context manager for handling IMAP connections"""
        imap = None
        try:
            print("Connecting to IMAP server...")
            imap = imaplib.IMAP4_SSL(self.imap_host, 993)
            imap.login(self.username, self.password)
            yield imap
        finally:
            if imap:
                try:
                    imap.close()
                except:
                    pass
                try:
                    imap.logout()
                except:
                    pass

    def decode_email_header(self, header: str) -> str:
        """Decode email header"""
        decoded_header = decode_header(header)
        return ' '.join([
            text.decode(encoding or 'utf-8') if isinstance(text, bytes) else text
            for text, encoding in decoded_header
        ])

    def parse_sender_name(self, from_header: str) -> tuple[str, str]:
        """Extract first and last name from email sender"""
        # Remove email address part if present
        name_part = from_header.split('<')[0].strip().strip('"')

        # Split into first and last name
        name_parts = name_part.split()
        if len(name_parts) >= 2:
            return name_parts[0], ' '.join(name_parts[1:])
        elif len(name_parts) == 1:
            return name_parts[0], ""
        return "", ""

    def create_ticket_payload(self, msg: email.message.Message) -> Dict[str, Any]:
        """Create ticket payload from email message"""
        print("Creating ticket payload...")
        subject = self.decode_email_header(msg["subject"] or "No Subject")
        from_header = self.decode_email_header(msg["from"])
        sender_email = email.utils.parseaddr(from_header)[1]
        first_name, last_name = self.parse_sender_name(from_header)

        # Get email body
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode()
                    break
        else:
            body = msg.get_payload(decode=True).decode()

        return {
            "title": subject,
            "content": body,
            "priority": "Normal",
            "sender": {
                "first_name": first_name,
                "last_name": last_name,
                "email": sender_email
            }
        }

    def send_to_webhook(self, payload: Dict[str, Any]) -> bool:
        """Send ticket data to Fluent Support webhook"""
        try:
            response = requests.post(
                self.webhook_url,
                json=payload,
                headers={"Content-Type": "application/json"}
            )
            return response.status_code == 200
        except Exception as e:
            print(f"Error sending to webhook: {e}")
            return False

    def process_emails(self):
        """Process unread emails and create tickets"""
        with self.imap_connection() as imap:
            try:
                # Select inbox
                status, messages = imap.select("INBOX")
                if status != 'OK':
                    print("Failed to select INBOX")
                    return

                # Search for unread emails
                status, message_numbers = imap.search(None, "UNSEEN")
                if status != 'OK':
                    print("Failed to search for unread messages")
                    return

                # Check if we have any messages
                if not message_numbers or not message_numbers[0]:
                    print("No unread messages found")
                    return

                # Process each message
                for num in message_numbers[0].split():
                    try:
                        # Fetch email message
                        status, msg_data = imap.fetch(num, "(RFC822)")
                        if status != 'OK':
                            print(f"Failed to fetch message {num}")
                            continue

                        if not msg_data or not msg_data[0]:
                            print(f"No data received for message {num}")
                            continue

                        email_body = msg_data[0][1]
                        msg = email.message_from_bytes(email_body)

                        # Skip blocked senders
                        sender_email = email.utils.parseaddr(msg["from"])[1]
                        if sender_email in self.blocked_senders:
                            print(f"Skipping blocked sender: {sender_email}")
                            continue

                        # Create and send ticket
                        ticket_payload = self.create_ticket_payload(msg)
                        if self.send_to_webhook(ticket_payload):
                            # Mark email as read only if ticket creation was successful
                            imap.store(num, '+FLAGS', '\\Seen')
                            print(f"Created ticket for email: {ticket_payload['title']}")
                        else:
                            print(f"Failed to create ticket for email: {ticket_payload['title']}")

                    except Exception as e:
                        print(f"Error processing message {num}: {e}")
                        continue

            except Exception as e:
                print(f"Error during email processing: {e}")

            finally:
                # Cleanup
                try:
                    imap.close()
                except:
                    pass
                try:
                    imap.logout()
                except:
                    pass

config = {
        "imap_host"   : os.getenv("IMAP_HOST"),
        "username"    : os.getenv("USERNAME"),
        "password"    : os.getenv("PASSWORD"),
        "webhook_url" : os.getenv("WEBHOOK_URL")
    }

app = Flask(__name__)

@app.route('/api/run-python')
def run_script():
    parser = EmailTicketParser(**config)
    parser.process_emails()
    result = "Execution finished"
    return jsonify(result)

if __name__ == '__main__':
    app.run()
