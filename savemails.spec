# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['src\\savemails.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MailBackup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='config\\MailBackup.ico',
)

import os, shutil, json
configPath = os.path.join(DISTPATH, 'config')
shutil.rmtree(configPath, ignore_errors=True)
os.makedirs(configPath)
shutil.copyfile('./config/template.html', os.path.join(configPath, 'template.html'))
resources : dict = {
    "accounts" : [
        {
            "username" : "email@address.de",
            "port": 993,
            "server" : "imap.address.de"
        },
        {
            "username" : "email@address.de",
            "port": 993,
            "server" : "imap.address.de"
        }
    ],
    "storageLocation" : "./BACKUP"
}

with open(os.path.join( configPath, 'resources.json'), 'w') as resourceFile:
    resourceFile.write(json.dumps( resources, sort_keys=True, indent=4))

