
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Explicitly gather moviepy modules
added_files = []
moviepy_modules = [
    'moviepy', 'moviepy.editor', 'moviepy.config', 
    'moviepy.tools', 'moviepy.audio', 'moviepy.video',
    'moviepy.audio.fx', 'moviepy.audio.fx.all', 'moviepy.audio.io',
    'moviepy.video.fx', 'moviepy.video.fx.all', 'moviepy.video.io',
    'moviepy.audio.io.readers', 'moviepy.video.io.ffmpeg_reader',
    'moviepy.video.io.html_tools', 'moviepy.video.io.ffmpeg_writer'
]

a = Analysis(
    ['app.py', 'import_helper.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=moviepy_modules + [
        'PIL', 'PIL._imagingft', 'PIL.ImageFont', 'PIL.ImageDraw',
        'numpy', 'scipy', 'scipy.io', 'scipy.signal',
        'pyttsx3', 'pyttsx3.drivers', 'pyttsx3.drivers.sapi5',
        'win32com', 'win32com.client',
        'websocket', 'ssl', 'wave', 'hmac',
        'hashlib', 'urllib', 'urllib.parse', 'base64', 'datetime'
    ],  # 移除了gtts
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PPT转视频工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
