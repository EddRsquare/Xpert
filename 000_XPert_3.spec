# -*- mode: python ; coding: utf-8 -*-
import sys
sys.setrecursionlimit(sys.getrecursionlimit() * 5)

a = Analysis(
    ['000_XPert_3.py'],
    pathex=[],
    binaries=[],
    datas=[    ('\\\\Akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosEntrada\\000_dataDirectory.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\Akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosEntrada\\000_dataBasePDF.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\Akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosSalida\\001_AV_PAT_BBDD_Vertical_SITLAB.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\Akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosSalida\\006_AV_PAT_BBDD_Vertical_SEPE.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\Akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosSalida\\008_AV_PAT_BBDD_Vertical_INSS.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosSalida\\005_AV_PAT_BBDD_Vertical_DGT.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosEntrada\\000_dataPDFImageFlag.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\Akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosSalida\\001_AV_PAT_BBDD_Vertical_190.xlsx', '.'),  # Copia en el directorio raíz
    # Los archivos con 'fecha_actual' deben manejarse en tiempo de ejecución
    ('\\\\akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosSalida\\HisResultadoCalificaciones\\300_Resultado_calificacion_' + fecha_actual + '.xlsx', '.'),  # Copia en el directorio raíz
    ('\\\\akes-db-01-004\\SharedReports\\LegalPDFExtract\\FicherosSalida\\HisResultadoCalificaciones\\300_Resultado_calificacion_' + fecha_actual + '.csv', '.'),  # Copia en el directorio raíz
],
    hiddenimports=['pandas', 'openpyxl', 'tkinter', 'datetime'],
    hookspath=['C:/ProgramData/anaconda3/Scripts/my_hooks'],
    hooksconfig={},
    excludes=['sphinx', 'sphinxcontrib', 'sphinx.ext'],
    runtime_hooks=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='000_XPert_3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)