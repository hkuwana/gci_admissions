# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['GCI_merger.py'],
             pathex=['/Users/hirokuwana/SummerProjects2020/GCI'],
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
          [],
          exclude_binaries=True,
          name='GCI_merger',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='GCI_merger')
app = BUNDLE(coll,
             name='GCI_merger.app',
             icon=None,
             bundle_identifier=None)
