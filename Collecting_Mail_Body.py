import win32com.client
import os
import sys
import time

from Collecting_Details_Without_Duplicates import Searching_strings

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

f = open("Body.txt", "w+")


def Body():
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the index of a folder - say inbox
    messages = inbox.Items

    message = messages.GetLast()

    for m in messages:
        if m.Subject == 'XXXXX-XXXXXXXXX .*':
            ol = win32com.client.Dispatch("Outlook.Application")
            inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
            for message in inbox.Items:
                if message.UnRead == True:
                    print(message.body)
            Searching_strings()
            time.sleep(30)

        elif m.Subject == 'XXXXX-XXXXXXXXX .*':
            ol = win32com.client.Dispatch("Outlook.Application")
            inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
            for message in inbox.Items:
                if message.UnRead == True:
                    print(message.body)
            Searching_strings()
            time.sleep(30)

        elif m.SenderEmailAddress == 'sample@mail.com':
            if m.Subject == 'xxxxxx-xxxxxxx .*':
                ol = win32com.client.Dispatch("Outlook.Application")
                inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
                for message in inbox.Items:
                        if message.UnRead:
                            print(message.body)
                Searching_strings()
                time.sleep(30)

    exit()


if __name__ == "__main__":
    while True:
        Body()
        time.sleep(300)

#Searching_strings()
print('----------Successfull---------------')
