block_cipher = None


a = Analysis(['fera.py'],
             pathex=['/media/sf_B_DRIVE/VISUALIZADOR/FERA'],
             binaries=[],
             datas=[('./Imagens/fts4tutorial.png', '.'), ('./ListasDeBusca/*', './ListasDeBusca'), ('./Imagens/logoMini.ico', '.')],
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
               name='FERA_LINUX')
