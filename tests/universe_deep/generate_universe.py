"""非 Office 全宇宙加深样例生成（B–G，对照 FULL_UNIVERSE 大纲）。"""

from __future__ import annotations

import base64
import io
import json
import shutil
import struct
import wave
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
SAMPLES = ROOT / "samples"
ASSETS = ROOT / "assets"
FFMPEG = Path(__file__).resolve().parents[2] / "tools" / "ffmpeg" / "ffmpeg.exe"
PDF_SRC = Path(__file__).resolve().parents[1] / "office_pdf_wave4" / "samples"


def _font(size=16):
    for n in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\arial.ttf"):
        if Path(n).exists():
            try:
                return ImageFont.truetype(n, size)
            except OSError:
                pass
    return ImageFont.load_default()


def make_assets() -> dict[str, Path]:
    ASSETS.mkdir(parents=True, exist_ok=True)
    out = {}
    # 有字截图
    p = ASSETS / "text_shot.png"
    im = Image.new("RGB", (480, 120), (255, 255, 255))
    dr = ImageDraw.Draw(im)
    dr.text((20, 30), "UNIV_IMG_TEXT 华北增长12%", fill=(20, 20, 20), font=_font(22))
    dr.text((20, 70), "North China Revenue", fill=(37, 99, 235), font=_font(18))
    im.save(p)
    out["text"] = p

    # 表格截图（无 OCR 预期抽不出字）
    p = ASSETS / "table_shot.png"
    im = Image.new("RGB", (400, 160), (250, 250, 250))
    dr = ImageDraw.Draw(im)
    dr.rectangle([10, 10, 390, 150], outline=(0, 0, 0))
    dr.text((30, 40), "区 | 值", fill=(0, 0, 0), font=_font(16))
    dr.text((30, 80), "华北 | 1280", fill=(0, 0, 0), font=_font(16))
    im.save(p)
    out["table"] = p

    # 风景无字
    p = ASSETS / "scenic.png"
    Image.new("RGB", (320, 200), (70, 130, 180)).save(p)
    out["scenic"] = p

    # 小图标 / 大图 / 竖图
    Image.new("RGB", (32, 32), (220, 38, 38)).save(ASSETS / "icon.png")
    out["icon"] = ASSETS / "icon.png"
    Image.new("RGB", (1200, 800), (5, 150, 105)).save(ASSETS / "large.png")
    out["large"] = ASSETS / "large.png"
    Image.new("RGB", (400, 900), (124, 58, 237)).save(ASSETS / "portrait.png")
    out["portrait"] = ASSETS / "portrait.png"

    # JPEG + 简单 EXIF 感（Pillow 写 jpeg）
    p = ASSETS / "exif.jpg"
    im = Image.new("RGB", (400, 300), (240, 240, 240))
    ImageDraw.Draw(im).text((20, 140), "EXIF_JPEG_MARK", fill=(0, 0, 0), font=_font(20))
    im.save(p, "JPEG", quality=90)
    out["jpg"] = p

    # EPUB/HTML 用小图
    p = ASSETS / "epub_img.png"
    im = Image.new("RGB", (160, 60), (37, 99, 235))
    ImageDraw.Draw(im).text((10, 20), "EPUB_IMG", fill=(255, 255, 255), font=_font(14))
    im.save(p)
    out["epub"] = p
    return out


# ---------- PDF：复用 + 补缺 ----------
def gen_pdf() -> list[str]:
    names = []
    SAMPLES.mkdir(parents=True, exist_ok=True)
    wanted = [
        "PDF01_text_only.pdf",
        "PDF02_text_table.pdf",
        "PDF04_text_image_table.pdf",
        "PDF05_multi_table.pdf",
        "PDF07_multipage_sections.pdf",
        "PDF08_mixed_lang.pdf",
        "PDF09_two_column.pdf",
        "PDF10_stress_pack.pdf",
        "PDF14_scanned_like.pdf",
        "PDF18_encrypted.pdf",
        "PDF19_many_pages.pdf",
    ]
    for w in wanted:
        src = PDF_SRC / w
        if src.exists():
            dest = SAMPLES / f"U_{w}"
            shutil.copy2(src, dest)
            names.append(dest.name)
    # 表单域简易 PDF（AcroForm 可能抽不出，记行为）
    try:
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        pdfmetrics.registerFont(TTFont("MSYH", r"C:\Windows\Fonts\msyh.ttc"))
        p = SAMPLES / "U_PDF_form_fields.pdf"
        c = canvas.Canvas(str(p))
        c.setFont("MSYH", 14)
        c.drawString(72, 750, "PDF表单 PDF_FORM_MARK")
        c.drawString(72, 720, "姓名标签 NAME_LABEL")
        c.acroForm.textfield(name="name", tooltip="Name", x=72, y=690, width=200, height=20, borderWidth=1, borderColor=None, fillColor=None, textColor=None, forceBorder=True)
        c.save()
        names.append(p.name)
    except Exception as e:
        print("form pdf skip", e)
    return names


# ---------- EPUB ----------
def _write_epub(path: Path, chapters: list[tuple[str, str]], css: str = "", images: dict[str, Path] | None = None):
    """最小 EPUB3。chapters: (filename, html_body_inner)"""
    images = images or {}
    container = """<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>"""
    manifest_items = [
        '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
        '<item id="css" href="style.css" media-type="text/css"/>',
    ]
    spine = []
    nav_lis = []
    for i, (fn, _) in enumerate(chapters, 1):
        mid = f"ch{i}"
        manifest_items.append(f'<item id="{mid}" href="{fn}" media-type="application/xhtml+xml"/>')
        spine.append(f'<itemref idref="{mid}"/>')
        nav_lis.append(f'<li><a href="{fn}">Chapter {i}</a></li>')
    for name, _p in images.items():
        manifest_items.append(f'<item id="{name}" href="images/{name}" media-type="image/png"/>')

    opf = f"""<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="uid" version="3.0">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="uid">univ-epub-001</dc:identifier>
    <dc:title>Universe EPUB Book UNIV_EPUB_TITLE</dc:title>
    <dc:language>zh</dc:language>
  </metadata>
  <manifest>
    {''.join(manifest_items)}
  </manifest>
  <spine>
    {''.join(spine)}
  </spine>
</package>"""
    nav = f"""<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head><title>nav</title></head>
<body>
<nav epub:type="toc"><ol>{''.join(nav_lis)}</ol></nav>
</body></html>"""

    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
        zf.writestr("META-INF/container.xml", container)
        zf.writestr("OEBPS/content.opf", opf)
        zf.writestr("OEBPS/nav.xhtml", nav)
        zf.writestr("OEBPS/style.css", css or "h1{color:#dc2626;font-size:28px} p{color:#333}")
        for fn, body in chapters:
            xhtml = f"""<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>{fn}</title><link rel="stylesheet" href="style.css"/></head>
<body>{body}</body></html>"""
            zf.writestr(f"OEBPS/{fn}", xhtml)
        for name, img in images.items():
            zf.write(img, f"OEBPS/images/{name}")


def gen_epub(assets) -> list[str]:
    names = []
    p = SAMPLES / "U_EPUB_01_chapters.epub"
    _write_epub(
        p,
        [
            ("ch1.xhtml", "<h1>第一章 EP01_CH1</h1><p>开篇 EP01_BODY1 中英 mixed</p>"),
            ("ch2.xhtml", "<h1>第二章 EP01_CH2</h1><p>内容 EP01_BODY2</p><ul><li>要点A EP01_L1</li><li>要点B</li></ul>"),
            ("ch3.xhtml", "<h1>第三章 EP01_CH3</h1><p>收尾 EP01_END</p>"),
        ],
    )
    names.append(p.name)

    p = SAMPLES / "U_EPUB_02_list_table.epub"
    _write_epub(
        p,
        [
            (
                "ch1.xhtml",
                "<h1>列表与表 EP02_TITLE</h1>"
                "<ol><li>步骤1 EP02_N1</li><li>步骤2 EP02_N2</li></ol>"
                "<table><tr><th>区域</th><th>值</th></tr>"
                "<tr><td>华北</td><td>1280</td></tr>"
                "<tr><td>海外</td><td>430</td></tr></table>"
                "<p>EP02_MARK</p>",
            )
        ],
    )
    names.append(p.name)

    p = SAMPLES / "U_EPUB_03_image.epub"
    _write_epub(
        p,
        [
            (
                "ch1.xhtml",
                "<h1>含图章 EP03_TITLE</h1>"
                '<p>见图 EP03_MARK</p><img src="images/cover.png" alt="EPUB_IMG_ALT"/>'
                "<p>图注 CAP_EP03</p>",
            )
        ],
        images={"cover.png": assets["epub"]},
    )
    names.append(p.name)

    p = SAMPLES / "U_EPUB_04_css.epub"
    _write_epub(
        p,
        [("ch1.xhtml", "<h1 class='big'>CSS装饰 EP04_TITLE</h1><p style='color:red'>正文 EP04_BODY</p>")],
        css="h1.big{color:#7c3aed;font-size:40px} p{font-size:18px;color:#dc2626}",
    )
    names.append(p.name)

    p = SAMPLES / "U_EPUB_05_mixed_lang.epub"
    _write_epub(
        p,
        [
            (
                "ch1.xhtml",
                "<h1>中英混排 EP05_TITLE</h1>"
                "<p>Book Title: Quarterly Report《季度报告》 EP05_MARK</p>"
                "<p>North China（华北）+12%</p>",
            )
        ],
    )
    names.append(p.name)
    return names


# ---------- IPYNB ----------
def gen_ipynb(assets) -> list[str]:
    names = []

    def save(name, cells):
        nb = {
            "nbformat": 4,
            "nbformat_minor": 5,
            "metadata": {"kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"}},
            "cells": cells,
        }
        p = SAMPLES / name
        p.write_text(json.dumps(nb, ensure_ascii=False), encoding="utf-8")
        names.append(name)

    save(
        "U_IPYNB_01_md_only.ipynb",
        [
            {
                "cell_type": "markdown",
                "metadata": {},
                "source": ["# IP01 标题 IP01_H1\n", "\n", "正文 IP01_BODY\n", "\n", "- 列表 IP01_L1\n"],
            }
        ],
    )

    save(
        "U_IPYNB_02_code_stdout.ipynb",
        [
            {
                "cell_type": "code",
                "execution_count": 1,
                "metadata": {},
                "source": ["print('IP02_STDOUT_MARK')\n", "x = 1 + 1\n", "print(x)\n"],
                "outputs": [
                    {
                        "output_type": "stream",
                        "name": "stdout",
                        "text": ["IP02_STDOUT_MARK\n", "2\n"],
                    }
                ],
            }
        ],
    )

    save(
        "U_IPYNB_03_html_table.ipynb",
        [
            {
                "cell_type": "code",
                "execution_count": 1,
                "metadata": {},
                "source": ["# html table out\n"],
                "outputs": [
                    {
                        "output_type": "display_data",
                        "metadata": {},
                        "data": {
                            "text/html": [
                                "<table><tr><th>区域</th><th>值</th></tr>"
                                "<tr><td>华北</td><td>1280</td></tr></table>"
                                "<p>IP03_HTML_MARK</p>"
                            ],
                            "text/plain": ["IP03_PLAIN"],
                        },
                    }
                ],
            }
        ],
    )

    # image output
    buf = io.BytesIO()
    Image.open(assets["epub"]).save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    save(
        "U_IPYNB_04_image_out.ipynb",
        [
            {
                "cell_type": "markdown",
                "metadata": {},
                "source": ["# IP04 图输出 IP04_TITLE\n"],
            },
            {
                "cell_type": "code",
                "execution_count": 1,
                "metadata": {},
                "source": ["# show image\n"],
                "outputs": [
                    {
                        "output_type": "display_data",
                        "metadata": {},
                        "data": {"image/png": b64, "text/plain": ["IP04_IMG_PLAIN"]},
                    }
                ],
            },
        ],
    )

    save(
        "U_IPYNB_05_mixed.ipynb",
        [
            {"cell_type": "markdown", "metadata": {}, "source": ["# IP05 混合 IP05_TITLE\n", "说明 IP05_MARK\n"]},
            {
                "cell_type": "code",
                "execution_count": 1,
                "metadata": {},
                "source": ["print('IP05_CODE')\n"],
                "outputs": [{"output_type": "stream", "name": "stdout", "text": ["IP05_CODE\n"]}],
            },
            {
                "cell_type": "code",
                "execution_count": 2,
                "metadata": {},
                "source": ["pass\n"],
                "outputs": [
                    {
                        "output_type": "display_data",
                        "metadata": {},
                        "data": {
                            "text/html": ["<table><tr><td>华北</td><td>100</td></tr></table>"],
                            "text/plain": ["tbl"],
                        },
                    }
                ],
            },
        ],
    )

    save(
        "U_IPYNB_06_empty_raw.ipynb",
        [
            {"cell_type": "markdown", "metadata": {}, "source": []},
            {"cell_type": "code", "execution_count": None, "metadata": {}, "source": [], "outputs": []},
            {"cell_type": "raw", "metadata": {}, "source": ["RAW_IP06_MARK\n"]},
            {"cell_type": "markdown", "metadata": {}, "source": ["收尾 IP06_END\n"]},
        ],
    )
    return names


# ---------- Images ----------
def gen_images(assets) -> list[str]:
    names = []
    mapping = {
        "U_IMG_01_text_shot.png": assets["text"],
        "U_IMG_02_table_shot.png": assets["table"],
        "U_IMG_04_scenic.png": assets["scenic"],
        "U_IMG_05_exif.jpg": assets["jpg"],
        "U_IMG_06_icon.png": assets["icon"],
        "U_IMG_06_large.png": assets["large"],
        "U_IMG_07_portrait.png": assets["portrait"],
    }
    for name, src in mapping.items():
        dest = SAMPLES / name
        shutil.copy2(src, dest)
        names.append(name)
    # 海报多栏感
    p = SAMPLES / "U_IMG_03_poster.png"
    im = Image.new("RGB", (600, 300), (255, 255, 255))
    dr = ImageDraw.Draw(im)
    dr.rectangle([10, 10, 290, 290], outline=(0, 0, 0))
    dr.rectangle([310, 10, 590, 290], outline=(0, 0, 0))
    dr.text((40, 120), "LEFT_POSTER", fill=(0, 0, 0), font=_font(18))
    dr.text((340, 120), "RIGHT_POSTER", fill=(0, 0, 0), font=_font(18))
    im.save(p)
    names.append(p.name)
    return names


# ---------- Audio / Video ----------
def _write_wav(path: Path, freq=440, seconds=0.4, amp=0.3, silent=False):
    rate = 16000
    n = int(rate * seconds)
    with wave.open(str(path), "w") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        frames = bytearray()
        for i in range(n):
            if silent:
                val = 0
            else:
                import math

                val = int(amp * 32767 * math.sin(2 * math.pi * freq * i / rate))
            frames += struct.pack("<h", val)
        w.writeframes(frames)


def gen_av() -> list[str]:
    names = []
    SAMPLES.mkdir(parents=True, exist_ok=True)
    wav = SAMPLES / "U_AUD_01_tone.wav"
    _write_wav(wav, freq=523, seconds=0.5)
    names.append(wav.name)

    silent = SAMPLES / "U_AUD_03_silent.wav"
    _write_wav(silent, seconds=0.3, silent=True)
    names.append(silent.name)

    noise = SAMPLES / "U_AUD_03_noise.wav"
    rate = 16000
    with wave.open(str(noise), "w") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        import random

        frames = b"".join(struct.pack("<h", random.randint(-8000, 8000)) for _ in range(rate // 3))
        w.writeframes(frames)
    names.append(noise.name)

    if FFMPEG.exists():
        import subprocess

        env = dict(**{**dict(__import__("os").environ), "PATH": str(FFMPEG.parent) + ";" + __import__("os").environ.get("PATH", "")})
        mp3 = SAMPLES / "U_AUD_04_tone.mp3"
        subprocess.run(
            [str(FFMPEG), "-y", "-i", str(wav), "-codec:a", "libmp3lame", str(mp3)],
            check=False,
            capture_output=True,
            env=env,
        )
        if mp3.exists():
            names.append(mp3.name)

        m4a = SAMPLES / "U_AUD_05_tone.m4a"
        subprocess.run(
            [str(FFMPEG), "-y", "-i", str(wav), "-c:a", "aac", str(m4a)],
            check=False,
            capture_output=True,
            env=env,
        )
        if m4a.exists():
            names.append(m4a.name)

        mp4_audio = SAMPLES / "U_VID_01_audio_only.mp4"
        subprocess.run(
            [
                str(FFMPEG),
                "-y",
                "-f",
                "lavfi",
                "-i",
                "color=c=black:s=320x240:d=1",
                "-i",
                str(wav),
                "-c:v",
                "libx264",
                "-c:a",
                "aac",
                "-shortest",
                str(mp4_audio),
            ],
            check=False,
            capture_output=True,
            env=env,
        )
        if mp4_audio.exists():
            names.append(mp4_audio.name)

        mp4_silent = SAMPLES / "U_VID_02_no_audio.mp4"
        subprocess.run(
            [
                str(FFMPEG),
                "-y",
                "-f",
                "lavfi",
                "-i",
                "color=c=blue:s=320x240:d=1",
                "-c:v",
                "libx264",
                "-an",
                str(mp4_silent),
            ],
            check=False,
            capture_output=True,
            env=env,
        )
        if mp4_silent.exists():
            names.append(mp4_silent.name)
    else:
        print("ffmpeg missing, skip mp3/m4a/mp4")

    return names


# ---------- HTML / feeds / data ----------
def gen_web_data() -> list[str]:
    names = []

    def w(name, content):
        p = SAMPLES / name
        p.write_text(content, encoding="utf-8")
        names.append(name)

    w(
        "U_HTML_01_semantic.html",
        """<!DOCTYPE html><html><head><meta charset="utf-8"><title>HTML01</title></head>
<body>
<h1>语义全文 HTML01_H1</h1>
<p>段落 HTML01_BODY 中英 mixed</p>
<ul><li>列表 HTML01_L1</li><li>列表2</li></ul>
<ol><li>步骤 HTML01_N1</li></ol>
<blockquote>引用 HTML01_QUOTE</blockquote>
<pre>code_block HTML01_PRE</pre>
<table><tr><th>区域</th><th>值</th></tr><tr><td>华北</td><td>1280</td></tr></table>
<a href="https://example.com/h01">链接 HTML01_LINK</a>
<p>HTML01_MARK</p>
</body></html>""",
    )

    w(
        "U_HTML_02_css_noise.html",
        """<!DOCTYPE html><html><head><meta charset="utf-8"><style>
h1{color:#7c3aed;font-size:48px}.x{background:yellow}
</style><title>HTML02</title></head>
<body><h1 class="x">CSS噪声 HTML02_TITLE</h1>
<p style="color:red;font-size:20px">正文 HTML02_BODY</p>
<script>document.title='SHOULD_IGNORE_SCRIPT'</script>
</body></html>""",
    )

    w(
        "U_HTML_03_complex_table.html",
        """<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
<h1>复杂表 HTML03_TITLE</h1>
<table>
<thead><tr><th colspan="2">汇总 MERGE_H</th><th>备注</th></tr>
<tr><th>区域</th><th>营收</th><th>说明</th></tr></thead>
<tbody>
<tr><td>华北</td><td>1280</td><td>稳</td></tr>
<tr><td>海外</td><td>430</td><td>压</td></tr>
</tbody></table>
<p>HTML03_MARK</p>
</body></html>""",
    )

    w(
        "U_HTML_04_multi_media.html",
        """<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
<h1>多表多图 HTML04_TITLE</h1>
<p>HTML04_MARK</p>
<img src="assets_placeholder_a.png" alt="ALT_HTML04_A"/>
<img src="assets_placeholder_b.png" alt="ALT_HTML04_B"/>
<table><tr><th>A</th><th>1</th></tr><tr><td>华北</td><td>10</td></tr></table>
<table><tr><th>B</th><th>2</th></tr><tr><td>海外</td><td>20</td></tr></table>
</body></html>""",
    )

    w(
        "U_HTML_05_nav_footer.html",
        """<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
<nav>导航噪声 NAV_NOISE_HTML05</nav>
<main><h1>正文区 HTML05_TITLE</h1><p>关键句 HTML05_MARK</p></main>
<footer>页脚噪声 FOOTER_NOISE_HTML05</footer>
</body></html>""",
    )

    w(
        "U_HTML_06_script_style.html",
        """<!DOCTYPE html><html><head><meta charset="utf-8">
<style>.hide{display:none}</style>
<script>var x='SCRIPT_SHOULD_NOT_APPEAR_HTML06';</script>
</head><body>
<h1>脚本样式 HTML06_TITLE</h1>
<p>可见 HTML06_MARK</p>
<div class="hide">隐藏块</div>
</body></html>""",
    )

    w(
        "U_HTML_07_entities.html",
        """<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
<h1>实体与中文 HTML07_TITLE</h1>
<p>&lt;标签&gt; &amp; &quot;引号&quot; — 【华北】 HTML07_MARK</p>
<p>全角：１２８０</p>
</body></html>""",
    )

    w(
        "U_RSS_01_multi.rss",
        """<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0"><channel>
<title>RSS Feed UNIV_RSS_TITLE</title>
<link>https://example.com/rss</link>
<description>测试源</description>
<item>
  <title>条目1 RSS_ITEM_1</title>
  <link>https://example.com/1</link>
  <description><![CDATA[<p>描述 <b>RSS_DESC_1</b> 华北</p>]]></description>
</item>
<item>
  <title>条目2 RSS_ITEM_2</title>
  <description>RSS_DESC_2 Overseas</description>
</item>
</channel></rss>""",
    )

    w(
        "U_RSS_02_cn.rss",
        """<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0"><channel>
<title>中文源 RSS_CN_TITLE</title>
<item><title>季度报告 RSS_CN_ITEM</title>
<description>华北增长12% RSS_CN_MARK</description></item>
</channel></rss>""",
    )

    w(
        "U_ATOM_01_multi.atom",
        """<?xml version="1.0" encoding="utf-8"?>
<feed xmlns="http://www.w3.org/2005/Atom">
  <title>Atom Feed ATOM_TITLE</title>
  <entry><title>Entry1 ATOM_E1</title><summary>摘要 ATOM_S1</summary></entry>
  <entry><title>Entry2 ATOM_E2</title><content type="html">&lt;p&gt;内容 ATOM_C2 华北&lt;/p&gt;</content></entry>
</feed>""",
    )

    w(
        "U_CSV_01_basic.csv",
        "区域,营收,同比\n华北,1280,+12%\n海外,430,-5%\nCSV01_MARK,1,0\n",
    )
    w(
        "U_CSV_02_escape.csv",
        '名称,备注\n"华北,集团","含""引号""与逗号"\n海外,普通\nCSV02_MARK,ok\n',
    )
    # BOM + 宽表
    p = SAMPLES / "U_CSV_03_bom_wide.csv"
    header = ",".join([f"列{i}" for i in range(1, 16)])
    row = ",".join(str(i) for i in range(1, 16))
    p.write_text("\ufeff" + header + "\n" + row + "\nCSV03_MARK,2,3,4,5,6,7,8,9,10,11,12,13,14,15\n", encoding="utf-8")
    names.append(p.name)

    w(
        "U_JSON_01_nested.json",
        json.dumps(
            {
                "title": "JSON01_TITLE",
                "meta": {"region": "华北", "mark": "JSON01_MARK"},
                "items": [{"name": "A", "value": 1280}, {"name": "Overseas", "value": 430}],
            },
            ensure_ascii=False,
            indent=2,
        ),
    )
    w(
        "U_JSON_02_array.json",
        json.dumps([{"id": i, "label": f"J2_{i}"} for i in range(1, 31)] + [{"id": 99, "label": "JSON02_MARK"}], ensure_ascii=False),
    )
    w(
        "U_JSONL_01.jsonl",
        "\n".join(
            [
                json.dumps({"i": 1, "t": "JSONL_1"}, ensure_ascii=False),
                json.dumps({"i": 2, "t": "JSONL_2", "区域": "华北"}, ensure_ascii=False),
                json.dumps({"i": 3, "t": "JSONL_MARK"}, ensure_ascii=False),
            ]
        )
        + "\n",
    )

    w(
        "U_XML_01_biz.xml",
        """<?xml version="1.0" encoding="UTF-8"?>
<report id="1">
  <title>XML01_TITLE</title>
  <section name="收入"><item region="华北" value="1280"/><item region="海外" value="430"/></section>
  <mark>XML01_MARK</mark>
</report>""",
    )
    w(
        "U_XML_02_ns.xml",
        """<?xml version="1.0" encoding="UTF-8"?>
<ns:doc xmlns:ns="http://example.com/ns">
  <ns:title>XML02_TITLE</ns:title>
  <ns:body>命名空间 XML02_MARK 华北</ns:body>
</ns:doc>""",
    )
    w(
        "U_TXT_01_manual.txt",
        """# 说明书结构 TXT01_TITLE

## 安装
步骤一 TXT01_N1
步骤二 TXT01_N2

## 说明
华北增长。TXT01_MARK
""",
    )
    w(
        "U_MD_01_fidelity.md",
        """# MD01 标题 MD01_H1

段落 MD01_BODY

- 列表 MD01_L1

| 区域 | 值 |
| --- | --- |
| 华北 | 1280 |

```python
print('MD01_CODE')
```

[链接](https://example.com/md01)

MD01_MARK
""",
    )
    return names


# ---------- ZIP ----------
def gen_zip(assets) -> list[str]:
    names = []
    # 准备成员文件
    members = SAMPLES / "_zip_members"
    members.mkdir(exist_ok=True)
    (members / "readme.md").write_text("# ZIP内MD ZIP_MD_MARK\n华北\n", encoding="utf-8")
    (members / "data.csv").write_text("区域,值\n华北,100\nZIP_CSV_MARK,1\n", encoding="utf-8")
    (members / "note.txt").write_text("文本 ZIP_TXT_MARK\n", encoding="utf-8")
    shutil.copy2(assets["epub"], members / "pic.png")

    # 简单 html
    (members / "page.html").write_text(
        "<html><body><h1>ZIP内HTML ZIP_HTML_MARK</h1></body></html>", encoding="utf-8"
    )

    p = SAMPLES / "U_ZIP_01_mixed.zip"
    with zipfile.ZipFile(p, "w") as zf:
        for f in members.iterdir():
            zf.write(f, f.name)
    names.append(p.name)

    p = SAMPLES / "U_ZIP_02_cn_names.zip"
    with zipfile.ZipFile(p, "w") as zf:
        zf.writestr("说明 中文.md", "# 中文成员 ZIP_CN_MARK\n内容\n")
        zf.writestr("数据.csv", "k,v\n华北,1\n")
    names.append(p.name)

    p = SAMPLES / "U_ZIP_03_unsupported.zip"
    with zipfile.ZipFile(p, "w") as zf:
        zf.writestr("ok.txt", "ZIP_OK_MARK\n")
        zf.writestr("bin.xyz", b"\x00\x01UNSUPPORTED_BIN")
        zf.writestr("readme.md", "# ZIP_SKIP_MARK\n")
    names.append(p.name)

    # 嵌套 zip
    inner = io.BytesIO()
    with zipfile.ZipFile(inner, "w") as zf:
        zf.writestr("inner.txt", "ZIP_INNER_MARK\n")
    p = SAMPLES / "U_ZIP_04_nested.zip"
    with zipfile.ZipFile(p, "w") as zf:
        zf.writestr("outer.txt", "ZIP_OUTER_MARK\n")
        zf.writestr("inner.zip", inner.getvalue())
    names.append(p.name)

    # 压力包：多格式
    p = SAMPLES / "U_ZIP_05_stress.zip"
    with zipfile.ZipFile(p, "w") as zf:
        zf.writestr("a.md", "# ZIP_STRESS_A\n")
        zf.writestr("b.csv", "x,y\n1,2\nZIP_STRESS_B,0\n")
        zf.writestr("c.json", json.dumps({"m": "ZIP_STRESS_C"}, ensure_ascii=False))
        zf.writestr("d.html", "<h1>ZIP_STRESS_D</h1>")
        zf.write(assets["text"], "e.png")
    names.append(p.name)
    return names


# ---------- negatives ----------
def gen_negatives() -> list[str]:
    names = []
    (SAMPLES / "U_G_01_empty.txt").write_bytes(b"")
    names.append("U_G_01_empty.txt")
    (SAMPLES / "U_G_05_corrupt.html").write_text("<html><body><h1>未闭合", encoding="utf-8")
    names.append("U_G_05_corrupt.html")
    (SAMPLES / "U_G_05_bad.zip").write_bytes(b"PK\x03\x04" + b"\x00" * 20 + b"BADZIP")
    names.append("U_G_05_bad.zip")
    (SAMPLES / "U_G_03 中文 空格.json").write_text(
        json.dumps({"mark": "G03_JSON_MARK", "区域": "华北"}, ensure_ascii=False), encoding="utf-8"
    )
    names.append("U_G_03 中文 空格.json")
    # 后缀不符：html 内容命名为 .pdf
    (SAMPLES / "U_G_04_html_as.pdf").write_text(
        "<html><body><p>MISMATCH_HTML_AS_PDF</p></body></html>", encoding="utf-8"
    )
    names.append("U_G_04_html_as.pdf")
    (SAMPLES / "U_G_bmp.bmp").write_bytes(b"BM" + b"\x00" * 60)
    names.append("U_G_bmp.bmp")
    return names


def main():
    assets = make_assets()
    parts = []
    parts += gen_pdf()
    parts += gen_epub(assets)
    parts += gen_ipynb(assets)
    parts += gen_images(assets)
    parts += gen_av()
    parts += gen_web_data()
    parts += gen_zip(assets)
    parts += gen_negatives()
    # dedupe preserve order
    seen = set()
    uniq = []
    for n in parts:
        if n and n not in seen:
            seen.add(n)
            uniq.append(n)
    print(f"generated={len(uniq)}")
    for n in uniq:
        print(" ", n)


if __name__ == "__main__":
    main()
