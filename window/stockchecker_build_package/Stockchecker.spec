# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all, copy_metadata

block_cipher = None

# Collect Streamlit and related packages that PyInstaller often misses
streamlit_datas, streamlit_binaries, streamlit_hiddenimports = collect_all('streamlit')
pandas_datas, pandas_binaries, pandas_hiddenimports = collect_all('pandas')
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all('openpyxl')
gspread_datas, gspread_binaries, gspread_hiddenimports = collect_all('gspread')
gauth_datas, gauth_binaries, gauth_hiddenimports = collect_all('google.auth')

added_files = [
    ('Stockchecker.py', '.'),
]
if os.path.isdir('data'):
    added_files.append(('data', 'data'))

hiddenimports = []
hiddenimports += streamlit_hiddenimports
hiddenimports += pandas_hiddenimports
hiddenimports += openpyxl_hiddenimports
hiddenimports += gspread_hiddenimports
hiddenimports += gauth_hiddenimports
hiddenimports += [
    'streamlit.web.cli',
    'streamlit.runtime.scriptrunner.magic_funcs',
    'google_auth_oauthlib',
    'google.oauth2.service_account',
]

datas = []
datas += streamlit_datas + pandas_datas + openpyxl_datas + gspread_datas + gauth_datas
datas += added_files
try:
    datas += copy_metadata('streamlit')
    datas += copy_metadata('pandas')
    datas += copy_metadata('openpyxl')
    datas += copy_metadata('gspread')
    datas += copy_metadata('google-auth')
except Exception:
    pass

binaries = []
binaries += streamlit_binaries + pandas_binaries + openpyxl_binaries + gspread_binaries + gauth_binaries

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name='Stockchecker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Stockchecker',
)
