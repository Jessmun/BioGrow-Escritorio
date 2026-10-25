# -*- mode: python ; coding: utf-8 -*-
# Compilar con: pyinstaller BioGrow_Analytics.spec
from PyInstaller.utils.hooks import collect_dynamic_libs, collect_submodules

# mysql.connector importa por nombre, en tiempo de ejecución, sus plugins de
# autenticación (mysql.connector.plugins.*) y sus mensajes de error
# (mysql.connector.locales.*); PyInstaller no los detecta por sí solo.
mysql_hiddenimports = (
    collect_submodules('mysql.connector.plugins')
    + collect_submodules('mysql.connector.locales')
)

# DLLs de mysql/vendor/ (OpenSSL, SASL y plugins de autenticación) que usa
# la extensión C (_mysql_connector) desde la ruta relativa mysql/vendor/.
mysql_binaries = collect_dynamic_libs('mysql')


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=mysql_binaries,
    datas=[],
    hiddenimports=mysql_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='BioGrow_Analytics',
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
)

mysql_hiddenimports = (
    collect_submodules('mysql.connector.plugins')
    + collect_submodules('mysql.connector.locales')
    + ['scipy.stats', 'scipy.special._cython_special']
)