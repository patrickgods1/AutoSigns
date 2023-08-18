# -*- mode: python -*-

block_cipher = None


a = Analysis(['AutoSigns.py'],
             pathex=[],
             binaries=[],
             datas=[("Template-GBC.docx", "."),("Template-GBC.pptx", "."), ("Template-SFC.docx", "."),("Template-SFC.pptx", "."), ("README.txt", ".")],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='AutoSigns v0.9.1',
          debug=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='AutoSigns v0.9.1')
