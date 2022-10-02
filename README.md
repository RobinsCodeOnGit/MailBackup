# MailBackup

A tool to backup all mails from email accounts specified in `./config/resources.json` to a local folder in eml format. An index.html for all emails is also generated for ease of use and searching later on.

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
