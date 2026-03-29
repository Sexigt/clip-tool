# ── PyInstaller spec for ClipTool ──
# Usage: pyinstaller build.spec

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('audio/*.mp3', 'audio'),
    ],
    hiddenimports=[
        'bettercam',
        'cv2',
        'numpy',
        'pyaudiowpatch',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'pystray',
        'plyer',
        'plyer.platforms.win.notification',
        'faster_whisper',
        'customtkinter',
        'keyboard',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'pandas',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ClipTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
