# MailBackup

A tool to backup all mails from email accounts specified in `./config/resources.json` to a local folder in eml format. An index.html for all emails is also generated for ease of use and searching later on.
During backup progress a progressbar showing the number of folders and number of to be backed up mails are displayed using tkinter.

## Keywords
tkinter, progressBar, imaplib, mail, backup, eml-format

## Setup

On Windows:

### Create Environment

```cmd
git clone this-repository-url
cd MailBackup
python -m venv .
Scripts\activate
pip install -r requirements.txt
```

Sourcecode can be found in `./src`

### Create Executable with PyInstaller

```cmd
python -m PyInstaller ./savemails.spec
```

`MailBackup.exe` and `config/` can then be found in `./dist`.
Adjust `./dist/config/resources.json` before use.

## Terms of Use

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
