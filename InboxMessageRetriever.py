import win32com.client

class EmailDownloader(object):

    def retieveEmails(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        return outlook.GetDefaultFolder(6)

def messageRetriever(self):
        inbox = self.__retieveEmails
        messages = inbox.Items
        return messages.GetFirst()