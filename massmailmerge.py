import traceback
import argparse
import smtplib
import getpass
import os
import sys
import csv
import magic
from email.message import EmailMessage

parser = argparse.ArgumentParser(description="MassMailMerge. Sends mass mail to recipients list using message template with placeholders and attachments. Placeholders can be included in curly brackets. A placeholder's value for each recipient is extracted the column of the same name in the recipients file.")
parser.add_argument("--server", help="Email account SMTP host; will be prompted if omitted.")
parser.add_argument("--port", type=int, help="Email account SMTP port; will be prompted if omitted.")
parser.add_argument("--username", help="Email account username; will be prompted if omitted.")
parser.add_argument("--password", help="Email account password; will be prompted if omitted.")
parser.add_argument("--sender", help="Sender, e.g. \"Obi Wan <ow@anther.planet>\"; will be prompted if omitted.")
parser.add_argument("--subject", help="Subject line; may contain placeholders .")
parser.add_argument("--template", help="Message template file; may contain placeholders; will be prompted if omitted.")
parser.add_argument("--recipients",help="File containing list of recipients with at least column 'to' (recipient email address, see --emailcol) plus one column for any placeholder in the message template; will be prompted if omitted.")
parser.add_argument("--emailcol", default="to", help="Name of the column holding the email address of the recipient")
parser.add_argument("--delimiter", default="\t", help="Column delimiter in recipients file; default: \\t.")
parser.add_argument("--notls", default=False, action="store_true", help="Do not establish TLS connection to SMTP server.")
parser.add_argument("--cc", nargs="+", help="List of email addresses to include for carbon copy.")
parser.add_argument("--bcc", nargs="+", help="List of email addresses to include for bline carbon copy.")
parser.add_argument("--attach", nargs="+", help="List of column names in recipients file that should be treated as file paths to attach to the message; each argument here may contain placeholders.")
args = parser.parse_args()

sandbox = input("Sandbox run? (Y/n) ")
sandbox = sandbox.lower().strip() in ["", "y"]
smtpsrv = None

try:
    send_count = 0
    header = []

    if not sandbox:
        server = args.server if args.server else input("SMTP host: ")
        if not server: 
            raise Exception("Server missing.")
        
        port = args.port if args.port else int(input("SMTP port: "))
        if not port: 
            raise Exception("Port missing.")
        
        username = args.username if args.username else input("Username: ")
        if not username: 
            raise Exception("Username missing.")
        
        password = args.password if args.password else getpass.getpass("Password: ")
        if not password: 
            raise Exception("Password missing.")
        
    sender = args.sender if args.sender else input("Sender: ")
    if not sender: 
        raise Exception("Sender missing.")
    
    subject = args.subject if args.subject else input("Subject: ")
    if not subject: 
        raise Exception("Subject missing.")
    
    template = args.template if args.template else input("Message template file: ")
    if not template: 
        raise Exception("Message template file missing.")
    
    recipients = args.recipients if args.recipients else input("Recipients file: ")
    if not recipients: 
        raise Exception("Recipients file missing.")
    
    with open(template, "r", encoding="utf-8") as f, \
        open(recipients, "r", encoding="utf-8") as recipients_file: 
        
        template_text = f.read()
        
        if not sandbox:
            smtpsrv = smtplib.SMTP(server, port)
            if not args.notls:
                smtpsrv.starttls()
            smtpsrv.login(username, password)

        reader = csv.reader(recipients_file, delimiter=args.delimiter)        
        for row in reader:
            if len(header) == 0:
                header = row.copy()
                continue
                        
            row = {header[i]:row[i] for i in range(len(header))}
            email = EmailMessage()
            email.set_charset("utf-8")
            email.set_content(template_text.format(**row))
            email["Subject"] = subject.format(**row)
            email["From"] = sender
            email["To"] = row[args.emailcol]
            if args.cc:
                email["Cc"] = ",".join(args.cc)
            if args.bcc:
                email["Bcc"] = ",".join(args.bcc)
            if args.attach:
                for attachment in args.attach:
                    if not row[attachment]: # no file given
                        continue
                    filepath = row[attachment].format(**row)
                    if not os.path.exists(filepath):
                        print(f"WARNING: cannot attach '{filepath}', does not exist!", file=sys.stderr)
                        continue
                    with open(filepath, "rb") as f:
                        content = f.read()
                        mime = magic.from_buffer(content, mime=True)
                        mime = mime.split("/")
                        if not sandbox:
                            email.add_attachment(content, maintype=mime[0], subtype=mime[1], filename=os.path.basename(filepath))
            
            if not sandbox: 
                smtpsrv.send_message(email)
            
            print(f"Email sent to {row[args.emailcol]}")
            send_count += 1
    
    print(f"\n{send_count} message(s) sent.")
except Exception as e:
    print("ERROR!", e, mime)
    traceback.print_exc()
finally:
    if smtpsrv: 
        smtpsrv.quit()