# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

root = Path(SPECPATH).resolve()

a = Analysis(
    [str(root / 'rt_quic_editor.py')],
    pathex=[str(root)],
    binaries=[],
    datas=[
        (str(root / 'build_viewer.py'), '.'),
        (str(root / 'plotly.min.js'), '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='RT-QuIC-Editor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='RT-QuIC-Editor',
)

app = BUNDLE(
    coll,
    name='RT-QuIC-Editor.app',
    icon=str(root / 'assets' / 'appicon.icns'),
    bundle_identifier='com.rtquic.editor',
    info_plist={
        'CFBundleName': 'RT-QuIC-Editor',
        'CFBundleDisplayName': 'RT-QuIC Editor',
        'CFBundleShortVersionString': '2.2',
        'CFBundleVersion': '2.2',
        'NSHighResolutionCapable': True,
    },
)

