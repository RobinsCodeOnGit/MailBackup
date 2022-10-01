
from datetime import datetime
from email.message import Message
import imaplib
import email
from email.header import decode_header
import os
import base64
import json
import threading
from time import strftime
from tkinter import Label, StringVar, Tk, ttk, Button
import tkinter
import traceback

import jinja2


class ProgressWindow:
    def __init__(self, title: str) -> None:
        self.window = tkinter.Tk()
        self.window.title(title)
        self.progressBar = ttk.Progressbar(
            master=self.window, orient=tkinter.HORIZONTAL, length=300)
        self.progressBar.pack(padx=10, pady=5)
        self.messageVar = StringVar(self.window)
        self.messageLabel = Label(self.window, textvariable=self.messageVar)
        self.messageLabel.pack()
        self.button = None

    def addButtonWithFunction(self, buttonText: str, buttonFunction, parameters: tuple):
        # t = threading.Thread(target=self.__buttonClicked(
        #    buttonFunction, parameters))
        self.buttonFunction = buttonFunction
        self.buttonFunctionParameters = parameters
        self.button = Button(master=self.window,
                             text=buttonText, command=lambda: self.__buttonClicked())
        self.button.pack(pady=5)

    def __buttonClicked(self):
        if (self.button != None):
            self.button['state'] = tkinter.DISABLED
        threading.Thread(target=self.buttonFunction(
            self.buttonFunctionParameters)).start()

    def updateBar(self, message: str, percent: int):
        self.progressBar['value'] = percent
        self.messageVar.set(message)
        print(message + " " + str(percent) + "%")
        self.window.update()
        self.window.update_idletasks()

    def runGUI(self):
        print('Starting gui')
        self.window.mainloop()
        #threading.Thread(target=lambda: self.window.mainloop()).start()


class SaveMails:
    def __init__(self) -> None:
        self.resources = json.load(open('./config/resources.json'))
        self.imap = imaplib.IMAP4_SSL(self.resources['account']['server'])
        # authenticate
        self.imap.login(self.resources['account']['username'], self.decode(
            self.resources['account']['password_enc']))
        self.folders = [x.decode().split('"/" ')[1]
                        for x in self.imap.list()[1]]
        self.savedMails = []

    def createStorage(self):
        now = datetime.now().strftime('%Y%m%d-%H%M')
        self.backupFolder = os.path.join(
            self.resources['storageLocation'], self.resources['account']['username'], 'backup_' + now)
        if not os.path.exists(self.backupFolder):
            os.makedirs(self.backupFolder)
        self.folderLocations = []
        for folder in self.folders:
            cleanFoldername = SaveMails.clean(folder, '/\\').replace(
                '/', os.path.sep).replace('\\', os.path.sep)
            folderLocation = self.backupFolder + os.path.sep + cleanFoldername
            if not os.path.exists(folderLocation):
                os.makedirs(folderLocation)
            self.folderLocations.append(folderLocation)

    def backupMails(self, progress: ProgressWindow):
        for idx, folder in enumerate(self.folders):
            self.backupMailsInFolder(idx, folder, progress)

    def backupMailsInFolder(self, idx: int, folder: str, progress: ProgressWindow):
        status, countMessagesBytesList = self.imap.select(folder)
        if status != 'OK':
            print("Errro accessing folder: " + folder)
            return
        n = int(countMessagesBytesList[0])
        for i in range(1, n+1):
            res, raw_mail = self.imap.fetch(str(i), '(RFC822)')
            progressMessage = '({0}/{1}) Folder: {2} ({3}/{4}) - {5}'.format(
                idx, len(self.folders), folder, i, n, str(res))
            progressPercentage = int((float(i)/float(n))*100)
            progress.updateBar(progressMessage, progressPercentage)
            self.saveMail(
                raw_mail[0][1], self.folderLocations[idx], folderName=folder)

    def saveMail(self, rawMail: bytes, folderLocation, folderName):
        try:
            mail = email.message_from_bytes(rawMail)
            path = self.filepathOf(mail, folderLocation)
            with open(path, 'wb') as mailFile:
                mailFile.write(rawMail)
            self.savedMails.append({
                "folder": folderName,
                "fileLocation": path,
                "subject": self.subjectOf(mail),
                "from": mail.get('from'),
                "to": mail.get('to'),
                "date": email.utils.parsedate_to_datetime(mail.get('date')).strftime('%d.%m.%Y %H:%M Uhr'),
                "hasAttachment": mail.get('content-type').startswith('multipart/mixed;')
            })
            return True
        except Exception as e:
            traceback.print_exception(e)
            return False

    def filepathOf(self, mail: Message, folderLocation: str):
        dateStr = mail.get('date')
        date: datetime = email.utils.parsedate_to_datetime(dateStr)
        prefix = date.strftime('%Y%m%d-%H%M_')
        subject = SaveMails.clean(self.subjectOf(mail))
        path = os.path.join(folderLocation, prefix + subject + '.eml')
        while (os.path.exists(path)):
            path = path[:-4] + '(1).eml'
        return path

    def subjectOf(self, mail: Message):
        subject = mail.get('subject')
        if subject == None:
            return ''
        if type(subject) == str and not '=?' in subject:
            return subject
        try:
            decoded_pairs = email.header.decode_header(subject)
            subject = ''
            for pair in decoded_pairs:
                subjectBytes: bytes = pair[0]
                charset: str = pair[1]
                if charset == None:
                    charset = 'utf-8'
                try:
                    subject += subjectBytes.decode(charset)
                except:
                    try:
                        subject += subjectBytes.decode('ascii')
                    except Exception as exp:
                        print(
                            'Failed to extract Subject. Uknown encoding: ' + charset)
                        raise (exp)
        except Exception as e:
            traceback.print_exception(e)
        return str(subject)

    def clean(text, allowAdditionalCharacters=''):
        # clean text for creating a folder
        return "".join(c if c.isalnum() or c in '().-, =!' or c in allowAdditionalCharacters else " " for c in text)

    def encode(s: str):
        return base64.b64encode(s.encode(encoding='utf-8')).decode(encoding='utf-8')

    def decode(self, s: str):
        return base64.b64decode(s).decode(encoding='utf-8')

    def createHtml(self):
        templateLoader = jinja2.FileSystemLoader(searchpath="./config")
        templateEnv = jinja2.Environment(loader=templateLoader)
        TEMPLATE_FILE = "template.html"
        template = templateEnv.get_template(TEMPLATE_FILE)
        template.stream(account=self.resources['account']['username'], now=datetime.now().strftime(
            '%d.%m.%Y'), mails=self.savedMails).dump(os.path.join(self.backupFolder, 'index.html'))


def startBackup(parameters: tuple):
    mail: SaveMails = parameters[0]
    progress: ProgressWindow = parameters[1]
    mail.createStorage()
    progress.updateBar('Folder Structure created', 20)
    mail.backupMails(progress)
    mail.createHtml()


if __name__ == '__main__':
    mail = SaveMails()
    progressWindow = ProgressWindow('Mail Backup')
    progressWindow.addButtonWithFunction(
        'Backup Now', startBackup, (mail, progressWindow))
    progressWindow.runGUI()
