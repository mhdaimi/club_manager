# -*- mode: python -*-
a = Analysis(['address_book.py.py'],
             pathex=['C:\\Python27\\Lib\\site-packages\\PyInstaller-3.2'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
a.datas += [('logo.gif','C:\\Python27\\Lib\\site-packages\\PyInstaller-3.2\\logo.gif','DATA')]
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='address_book.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False , icon='C:\\Python27\\Lib\\site-packages\\PyInstaller-3.2\\logo.gif')