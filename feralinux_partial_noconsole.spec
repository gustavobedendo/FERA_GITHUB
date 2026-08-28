import os

block_cipher = None
project_root = os.path.abspath(SPECPATH)


a = Analysis(['fera_partial.py'],
             pathex=[project_root],
             binaries=[],
             datas=[('./ListasDeBusca/*', './ListasDeBusca'), ('./Imagens/*', './Imagens'), ('ffmpeg.exe','.'), ('whats_new.txt', '.')],
             hiddenimports=['PIL', 'PIL._imagingtk', 'PIL._tkinter_finder'],
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
          name='fera',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='./Imagens/logoMini.ico', version='version.rc')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='fera')
