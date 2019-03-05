# -*- mode: python -*-

block_cipher = None


a = Analysis(['topicmodel.py'],
             pathex=['./rsa-conference-2019'],
             binaries=[],
             datas=[],
             hiddenimports=['cymem', 'murmurhash'],
             hookspath=['./hooks'],
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
          name='TopicModel',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )

app = BUNDLE(exe,
         name='myscript.app',
         icon='sd.ico',
         bundle_identifier=None)