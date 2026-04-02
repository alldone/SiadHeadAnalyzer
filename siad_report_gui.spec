# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Siad1_struttura_v2024.xsd', '.'),
        ('Siad2_struttura_v2024 (1).xsd', '.'),
        ('specialistica_verticale/FORMAT_ACQUAVIVA.xlsx', 'specialistica_verticale'),
        ('specialistica_verticale/BRANCA_Codici regionali-Codici SSN.xlsx', 'specialistica_verticale'),
        ('specialistica_verticale/NOTE_ETL_SPECIALISTICA.md', 'specialistica_verticale'),
    ],
    hiddenimports=[
        'mobilita_verticale.mobilita_gui',
        'mobilita_verticale.mobilita_report',
        'specialistica_verticale.specialistica_gui',
        'specialistica_verticale.etl_bancadati',
        'specialistica_verticale.validate_output',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SiadHeadAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SiadHeadAnalyzer',
)
