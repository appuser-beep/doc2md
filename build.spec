# -*- mode: python ; coding: utf-8 -*-

"""PyInstaller 规格：GUI exe + CLI exe。"""



from PyInstaller.utils.hooks import collect_all, collect_submodules



datas = []

binaries = []

hiddenimports = []



for pkg in (

    "markitdown",

    "customtkinter",

    "magika",

    "mammoth",

    "pptx",

    "openpyxl",

    "pdfminer",

    "pdfplumber",

    "bs4",

    "charset_normalizer",

    "olefile",

    "pydub",

    "speech_recognition",

    "youtube_transcript_api",

    "defusedxml",

    "xlrd",

    "pandas",

):

    try:

        d, b, h = collect_all(pkg)

        datas += d

        binaries += b

        hiddenimports += h

    except Exception:

        pass



hiddenimports += collect_submodules("markitdown")

common_hidden = [

    "converter",

    "cleanup",

    "excel_convert",

    "ipynb_convert",

    "zip_convert",

    "llm_settings",

    "azure_settings",

    "advanced_settings",

    "plugin_loader",

    "cli",

    "openai",

    "markitdown_ocr",

    "customtkinter",

    "markitdown",

    "mammoth",

    "docx",

    "pptx",

    "openpyxl",

    "xlrd",

    "pdfminer",

    "pdfplumber",

    "olefile",

    "pydub",

    "speech_recognition",

    "youtube_transcript_api",

    "magika",

    "markdownify",

    "bs4",

    "azure",

    "azure.core",

    "azure.identity",

    "azure.ai.documentintelligence",

    "azure.ai.contentunderstanding",

]

hiddenimports += common_hidden



a_gui = Analysis(

    ["main.py"],

    pathex=[],

    binaries=binaries,

    datas=datas,

    hiddenimports=hiddenimports + ["app"],

    hookspath=[],

    hooksconfig={},

    runtime_hooks=[],

    excludes=[],

    noarchive=False,

    optimize=0,

)



a_cli = Analysis(

    ["cli.py"],

    pathex=[],

    binaries=binaries,

    datas=datas,

    hiddenimports=hiddenimports,

    hookspath=[],

    hooksconfig={},

    runtime_hooks=[],

    excludes=["customtkinter", "tkinter"],

    noarchive=False,

    optimize=0,

)



pyz_gui = PYZ(a_gui.pure)

pyz_cli = PYZ(a_cli.pure)



exe_gui = EXE(

    pyz_gui,

    a_gui.scripts,

    a_gui.binaries,

    a_gui.datas,

    [],

    name="文档转Markdown",

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



exe_cli = EXE(

    pyz_cli,

    a_cli.scripts,

    a_cli.binaries,

    a_cli.datas,

    [],

    name="doc2md-cli",

    debug=False,

    bootloader_ignore_signals=False,

    strip=False,

    upx=True,

    upx_exclude=[],

    runtime_tmpdir=None,

    console=True,

    disable_windowed_traceback=False,

    argv_emulation=False,

    target_arch=None,

    codesign_identity=None,

    entitlements_file=None,

)

