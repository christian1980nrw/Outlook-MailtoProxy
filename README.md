# Outlook-MailtoProxy
Proof of concept to convert a Mailto-Link by parsing it and forwarding it to Outlook 365 by VBScript code

I needed this because of the (compared to HCL Notes) limited features like possible multiple attachments of Outlook to create Mails by the old mailto interface nowadays.
With this little tool we were are able to migrate a legacy software that was using mailto-links from Notes to Outlook.

- place the MailtoProxy.vbs-file at c:\ and register it with the attached register_MailtoProxy.reg registry file.
- After that go to your settings and set "Microsoft Console Based Script Host" as default Email program for mailto Links
- The script will convert the Mailto-Link and open a new mail in Outlook (tested with Outlook 2019 & Outlook 365 at Windows 11)
- The script creates a logfile at %temp%/MailtoProcessingLog.txt and adds a signature to the email by reading the personal data from the Microsoft Active Directory
  
Test-Syntax at a cmd:

cscript.exe c:\MailtoProxy.vbs mailto:test@test.de

