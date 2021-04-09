import sys

sys.setrecursionlimit(5000)

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['자체점검자동입력.py'],
             pathex=['C:\\Users\\win10\\Desktop\\자동화 프로그램\\자체점검자동입력'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='자체점검자동입력',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
