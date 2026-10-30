# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller Specification for Question Categorizer Desktop App
-------------------------------------------------------------
This file instructs PyInstaller how to bundle:
1. The desktop launcher (run_desktop.py) as entry point.
2. The application files (app.py and pipeline.py).
3. The Streamlit web frontend static assets (HTML/JS/CSS).
4. All package metadata and dynamic hidden imports.
5. Custom branding:
   - Windows: icon.ico
   - macOS: icon.icns
6. Cross-platform targets:
   - Windows: QuestionCategorizer/ directory (containing QuestionCategorizer.exe)
   - macOS: QuestionCategorizer.app bundle
"""

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, copy_metadata, collect_submodules

block_cipher = None

# Determine base paths
spec_root = os.path.abspath(SPECPATH)

# Look for app.py and pipeline.py either in the same directory or in the sibling "Streamlit approach" directory
def locate_source(filename):
    local_path = os.path.join(spec_root, filename)
    if os.path.exists(local_path):
        return local_path
    sibling_path = os.path.join(spec_root, "..", "Streamlit approach", filename)
    if os.path.exists(sibling_path):
        return os.path.abspath(sibling_path)
    return filename

app_py = locate_source('app.py')
pipeline_py = locate_source('pipeline.py')

# Locate icon files
def locate_icon(icon_name):
    p = os.path.join(spec_root, icon_name)
    if os.path.exists(p):
        return p
    p_assets = os.path.join(spec_root, 'assets', icon_name)
    if os.path.exists(p_assets):
        return p_assets
    return None

ico_path = locate_icon('icon.ico')
icns_path = locate_icon('icon.icns')

def safe_copy_metadata(pkg):
    try:
        return copy_metadata(pkg)
    except Exception:
        return []

# Collect application datas
datas = [
    (app_py, '.'),
    (pipeline_py, '.'),
]

if ico_path:
    datas.append((ico_path, '.'))
if icns_path:
    datas.append((icns_path, '.'))

# Collect Streamlit static web assets and package metadata
datas += collect_data_files('streamlit')
datas += safe_copy_metadata('streamlit')
datas += safe_copy_metadata('google-genai')
datas += safe_copy_metadata('google-generativeai')
datas += safe_copy_metadata('tqdm')
datas += safe_copy_metadata('requests')
datas += safe_copy_metadata('packaging')

# Collect hidden imports needed by dynamic execution
hiddenimports = [
    'streamlit',
    'streamlit.web.cli',
    'streamlit.runtime.caching',
    'streamlit.runtime.scriptrunner.magic_expressions',
    'streamlit.elements',
    'streamlit.components.v1',
    'openpyxl',
    'openpyxl.cell._writer',
    'docx',
    'pypdf',
    'pdfplumber',
    'google.genai',
    'google.genai.types',
    'google.generativeai',
    'pandas',
    'pipeline',
]
hiddenimports += collect_submodules('streamlit')
hiddenimports += collect_submodules('pipeline')

a = Analysis(
    ['run_desktop.py'],
    pathex=[spec_root, os.path.dirname(app_py)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='QuestionCategorizer',
    icon=ico_path,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # Keeps helpful console visible
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='QuestionCategorizer',
)

# On macOS, generate a native .app application bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='QuestionCategorizer.app',
        icon=icns_path,
        bundle_identifier='com.oursilhouette.questioncategorizer',
        info_plist={
            'CFBundleName': 'Question Categorizer',
            'CFBundleDisplayName': 'Question Categorizer',
            'CFBundleGetInfoString': 'Competitive Exam Question Categorizer',
            'CFBundleIdentifier': 'com.oursilhouette.questioncategorizer',
            'CFBundleVersion': '1.0.0',
            'CFBundleShortVersionString': '1.0.0',
            'NSHighResolutionCapable': 'True',
        },
    )
