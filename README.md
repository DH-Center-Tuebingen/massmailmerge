# Mass Mail Merge

Python script to send mass emails to a list of recipients using a message template with optional placeholders and attachments via your mail provider's SMTP server using your user credentials. Placeholders can be included in the message template using curly brackets. Each placeholder's value for a recipient is extracted from the corresponding column in the recipients CSV file.

## Example

Suppose you want to send an email using the following `appointment.txt` template to a list of recipients:

```
Dear {name}!

Your appointment is confirmed for {time} tomorrow. Please complete the attached form and bring it with you.

Best regards,
Dr. Toothtwissle
```

You have a list of recipients in `recipients.csv`. The CSV file should have one column for each placeholder in the template, one for the recipient's email address, and optionally one or more columns for each file attachment. Example:

```
name,email,time,attachment
Zephyr Quixote,zq@some-mail.service,2:00 PM,form.pdf
Lyra Nebulon,ln@some-mail.service,2:20 PM,form.pdf
Thalor Vortex,tv@some-mail.service,2:40 PM,form.pdf
Elara Starwind,es@some-mail.service,3:00 PM,form.pdf
Draven Myst,dm@some-mail.service,3:20 PM,form.pdf
```

The columns `name` and `time` will be used to replace the `{name}` and `{time}` placeholders in the email template. The `email` column will be referenced in the command line when invoking the script as the recipient's email address. The `attachment` column references the file `form.pdf` to attach to each email message. The command to send the messages would be, for instance:

```bash
python massmailmerge.py --subject "Your appointment confirmed" --sender "Dr. Toothtwissle <dr@tooth.twissle>" --template template.txt --recipients recipients.csv --emailcol email --cc some@cc.emailaddress --attach attachment
```

The script will then interactively ask for the SMTP server address and port, as well as user credentials. These can also be provided as command line arguments.

Licensed under the [MIT License](./LICENSE).
