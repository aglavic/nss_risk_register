# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['update_database.py'],
             pathex=[],
             binaries=[],
             datas=[],
             hiddenimports=['openpyxl.cell._writer'],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=['matplotlib', 'PyQt5'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

EXE(pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='update_database',
    debug=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=True )

# exe = EXE(pyz,
          # a.scripts, 
          # [],
          # exclude_binaries=True,
          # name='update_database',
          # debug=False,
          # bootloader_ignore_signals=False,
          # strip=False,
          # upx=True,
          # console=True,
          # disable_windowed_traceback=False,
          # target_arch=None,
          # codesign_identity=None,
          # entitlements_file=None )
# coll = COLLECT(exe,
               # a.binaries,
               # a.zipfiles,
               # a.datas, 
               # strip=False,
               # upx=True,
               # upx_exclude=[],
               # name='updat_database')
