block_cipher = None


a = Analysis(['fera_partial.py'],	
             pathex=['B:\\VISUALIZADOR\\FERA'],
             binaries=[],
             datas=[('./ListasDeBusca/*', './ListasDeBusca'), ('./Imagens/*', './Imagens'), ('ffmpeg.exe','.'), ('whats_new.txt', '.'), ('./third_party/7zip/*', './7zip')],
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
          console=True , icon='./Imagens/logoMini.ico', version='version.rc')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='fera')
