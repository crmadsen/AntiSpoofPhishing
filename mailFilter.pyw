import win32com.client
import win32api
import pythoncom
import re
import whois

# set organization name used for screening and the domain used for internal emails
# only a portion of the organization used in the whois query needs to be used
orgFlag = 'Wombat'

class Handler_Class(object):

    def __init__(self):
        count = 0
        phish = False
        checkFlag = True
        print('Scanning inbox for unread emails...')

        # First action to do when using the class in the DispatchWithEvents     
        inbox = self.Application.GetNamespace("MAPI").GetDefaultFolder(6)
        junk = self.Application.GetNamespace("MAPI").GetDefaultFolder(23)
        messages = inbox.Items

        # Check for unread emails when starting the event
        for mail in messages:
            if mail.UnRead:
                count = count+1

                # differentiate Exchange from external emails and generate the sender's email address
                # Exchange emails will not result in a whois request as they're considered legitimate
                if mail.SenderEmailType == 'EX':
                    sender = mail.Sender.GetExchangeUser().PrimarySmtpAddress
                    checkFlag = False
                else:
                    sender = mail.SenderEmailAddress 

                # split the domain off from the email address and perform a whois lookup
                senderDomain = sender.split('@')[1]

                if checkFlag:
                    try:
                        who = whois.whois(senderDomain)
                    except:
                        who = 'placeholder'
                        pass
                else:
                    who = 'placeholder'

                # if Wombat is detected in the organization element, move the email to the junk folder and alert the user
                try: 
                    if orgFlag in who.org:
                        win32api.MessageBox(0, 'A phishing email from a domain owned by Wombat Security Technologies was detected and moved to your junk email box...', 'Phishing Email Detected', 0x00001000) 
                        mail.Move(junk)
                        phish = True
                    else:
                        phish = False
                except:
                    pass
        if phish:
            print('Inbox scan complete...\n\t',count,'unread email(s) found.\n\tA phishing attempt was detected and quarantined.')
        else:
            print('Inbox scan complete...\n\t',count,'unread email(s) found.\n\tNo phishing attempts detected.')

        print('\nBeginning live scan...\n')

    def OnNewMailEx(self, receivedItemsIDs):
        phish = False
        checkFlag = True

        # ReceivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            junk = self.Application.GetNamespace("MAPI").GetDefaultFolder(23)

            # differentiate Exchange from external emails and generate the sender's email address
            if mail.SenderEmailType == 'EX':
                sender = mail.Sender.GetExchangeUser().PrimarySmtpAddress
                checkFlag = False
            else:
                sender = mail.SenderEmailAddress 

            # split the domain off from the email address and perform a whois lookup
            senderDomain = sender.split('@')[1]
            if checkFlag:
                try:
                    who = whois.whois(senderDomain)
                except:
                    who = 'placeholder'
                    pass
            else:
                who = 'placeholder'

            # if Wombat is detected in the organization element, move the email to the junk folder and alert the user
            try: 
                if orgFlag in who.org:
                    win32api.MessageBox(0, 'A phishing email from a domain owned by Wombat Security Technologies was detected and moved to your junk email box...', 'Phishing Email Detected', 0x00001000) 
                    mail.Move(junk)
                    phish = True
                else:
                    phish = False
            except:
                pass

        if phish:
            print('Incoming email...\n\tA phishing attempt was detected and quarantined.')
        else:
            print('Incoming email...\n\tNo phishing attempts detected.')

outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)

# loop
pythoncom.PumpMessages() 