#!/usr/bin/env python3
"""
Folio — минималистичный текстовый редактор для Windows.
Открывает и сохраняет: .docx (Word), .txt, .md, .html. Экспорт в PDF и печать.
Проекты и папки (настоящие папки в «Документы\\Folio»), карточки-ссылки, ссылки в тексте.
"""
import sys
import os
import re
import io
import json
import html
import shutil
import zipfile
import zlib
import subprocess
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import (Qt, QSize, QUrl, QBuffer, QByteArray, QIODevice, QTimer,
                            QSettings, Signal, QLibraryInfo, QTranslator, QRectF, QMarginsF,
                            QStandardPaths, QMimeData, QFileSystemWatcher, QFile, QPoint, QRect,
                            QObject, QEvent, QDateTime, QPointF, QSizeF)
from PySide6.QtGui import (QCursor, QFontDatabase, QAction, QIcon, QPixmap, QPainter, QColor, QFont, QTextCharFormat,
                           QTextBlockFormat, QTextListFormat, QTextCursor, QTextDocument,
                           QKeySequence, QImage, QTextImageFormat, QTextTableFormat, QTextLength,
                           QTextFrameFormat, QTextFormat, QBrush, QGuiApplication, QShortcut, QPageSize,
                           QPageLayout, QDesktopServices, QPainterPath, QPen, QFontMetrics)
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QTextEdit, QToolButton,
                               QPushButton, QHBoxLayout, QVBoxLayout, QLabel, QComboBox,
                               QFontComboBox, QFrame, QFileDialog, QMessageBox, QMenu,
                               QWidgetAction, QGridLayout, QLineEdit, QDialog, QSpinBox,
                               QSizePolicy, QListWidget, QListWidgetItem, QStyledItemDelegate,
                               QStyle, QStackedWidget, QAbstractItemView, QListView, QToolTip,
                               QCheckBox, QColorDialog)
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtPrintSupport import QPrinter, QPrintDialog

APP_NAME = "Folio"
APP_VERSION = "1.0.0"
GITHUB_REPO = "Romasha22/Folio"
DEFAULT_FONT = "Segoe UI"
DEFAULT_SIZE = 12
HEADING_SIZES = {1: 24, 2: 18, 3: 15}
TEXT_EXTS = {".txt", ".log", ".csv", ".ini", ".json", ".xml", ".cfg", ".py", ".js", ".css", ".bat"}
MD_EXTS = {".md", ".markdown"}
HTML_EXTS = {".html", ".htm"}
DOC_EXTS = {".docx"} | TEXT_EXTS | MD_EXTS | HTML_EXTS
LINK_COLOR = "#3D6BFF"
PALETTE = ["#3D5AFE", "#7C4DFF", "#E5484D", "#F76B15", "#F5B400",
           "#30A46C", "#12A594", "#0090FF", "#D6409F", "#8D6E63"]
INTERNAL_MIME = "application/x-folio-paths"
MONTHS = ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]


def res(name):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


# ───────────────────────────── Темы ─────────────────────────────
THEMES = {
    "light": dict(bg="#F1F3F8", surface="#FFFFFF", page="#FFFFFF", border="#E2E6EE",
                  text="#1A1D24", muted="#6B7280", accent="#3D5AFE", accent_text="#FFFFFF",
                  soft="#E7EBFF", hover="#EEF1F6", danger="#E5484D", sel="#C9D3FF"),
    "dark": dict(bg="#0E1014", surface="#161920", page="#161920", border="#252A34",
                 text="#E6E9EF", muted="#8A93A3", accent="#7C8CFF", accent_text="#0E1014",
                 soft="#232A48", hover="#1E222B", danger="#FF6B6B", sel="#34407A"),
}

def mix(c1, c2, t):
    a, b = QColor(c1), QColor(c2)
    return QColor(round(a.red() + (b.red() - a.red()) * t), round(a.green() + (b.green() - a.green()) * t),
                  round(a.blue() + (b.blue() - a.blue()) * t)).name()


def luminance(c):
    c = QColor(c)

    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * ch(c.red()) + 0.7152 * ch(c.green()) + 0.0722 * ch(c.blue())


def build_theme(dark, accent=None, tint=False, page=None, text=None):
    th = dict(THEMES["dark" if dark else "light"])
    if accent:
        acc = QColor(accent)
        if dark and luminance(acc.name()) < 0.18:
            acc = acc.lighter(150)
        th["accent"] = acc.name()
    acc = th["accent"]
    th["accent_text"] = "#FFFFFF" if luminance(acc) < 0.42 else "#111318"
    th["soft"] = mix(th["surface"], acc, 0.22 if dark else 0.13)
    th["sel"] = mix(th["page"], acc, 0.40 if dark else 0.26)
    if tint:
        th["bg"] = mix(th["bg"], acc, 0.08 if dark else 0.07)
        th["hover"] = mix(th["hover"], acc, 0.08)
        th["border"] = mix(th["border"], acc, 0.12)
        th["surface"] = mix(th["surface"], acc, 0.03 if dark else 0.015)
    if page:
        th["page"] = page
    if text:
        th["doc_text"] = text
    elif page:
        th["doc_text"] = "#1A1D24" if luminance(page) > 0.35 else "#E6E9EF"
    else:
        th["doc_text"] = th["text"]
    return th


# ───────────────────────────── Иконки ─────────────────────────────
SVG = {
    "new": '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M12 11v6M9 14h6"/>',
    "open": '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    "undo": '<path d="M9 14L4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
    "redo": '<path d="M15 14l5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
    "find": '<circle cx="11" cy="11" r="6.5"/><path d="M20 20l-4.2-4.2"/>',
    "more": '<circle cx="12" cy="5.5" r="1.4" fill="{c}"/><circle cx="12" cy="12" r="1.4" fill="{c}"/><circle cx="12" cy="18.5" r="1.4" fill="{c}"/>',
    "moon": '<path d="M20 14.5A8.5 8.5 0 1 1 9.5 4a6.8 6.8 0 0 0 10.5 10.5z"/>',
    "sun": '<circle cx="12" cy="12" r="4"/><path d="M12 2.5v2M12 19.5v2M2.5 12h2M19.5 12h2M5.3 5.3l1.4 1.4M17.3 17.3l1.4 1.4M5.3 18.7l1.4-1.4M17.3 6.7l1.4-1.4"/>',
    "bullets": '<circle cx="5" cy="7" r="1.3" fill="{c}"/><circle cx="5" cy="12" r="1.3" fill="{c}"/><circle cx="5" cy="17" r="1.3" fill="{c}"/><path d="M9.5 7H20M9.5 12H20M9.5 17H20"/>',
    "numbers": '<path d="M4 5.5l1.5-1v5"/><path d="M3.8 14.2a1.5 1.5 0 1 1 2.4 1.6L3.8 19h2.8"/><path d="M10 7H20M10 12H20M10 17H20"/>',
    "left": '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>',
    "center": '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>',
    "right": '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
    "justify": '<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>',
    "image": '<rect x="3" y="4.5" width="18" height="15" rx="2.5"/><circle cx="9" cy="10" r="1.8"/><path d="M21 15.5l-4.5-4.5L8 19.5"/>',
    "table": '<rect x="3" y="4.5" width="18" height="15" rx="2.5"/><path d="M3 9.5h18M3 14.5h18M9.5 9.5v10"/>',
    "marker": '<path d="M14.5 4.5l5 5-8 8H6.5v-5z"/><path d="M4 20.5h16" stroke-width="2.6"/>',
    "close": '<path d="M6 6l12 12M18 6L6 18"/>',
    "up": '<path d="M6 15l6-6 6 6"/>',
    "down": '<path d="M6 9l6 6 6-6"/>',
    "chev": '<path d="M7 10l5 5 5-5"/>',
    "back": '<path d="M15 5l-7 7 7 7"/>',
    "plus": '<path d="M12 5v14M5 12h14"/>',
    "link": '<path d="M10 14a4.5 4.5 0 0 0 6.4 0l3-3a4.5 4.5 0 0 0-6.4-6.4l-1.2 1.2"/><path d="M14 10a4.5 4.5 0 0 0-6.4 0l-3 3a4.5 4.5 0 0 0 6.4 6.4l1.2-1.2"/>',
    "globe": '<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"/>',
    "folder": '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    "file": '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
    "tune": '<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/><circle cx="15" cy="7" r="2.2"/><circle cx="9" cy="17" r="2.2"/>',
    "export": '<path d="M12 15V3M7.5 7.5L12 3l4.5 4.5"/><path d="M5 13v5a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-5"/>',
    "crop": '<path d="M6 2.5V16a2 2 0 0 0 2 2h13.5"/><path d="M2.5 6H16a2 2 0 0 1 2 2v13.5"/>',
    "arrow": '<path d="M4 18c4-9 9-12 15-12"/><path d="M14 4.5l5 1.5-1.5 5"/>',
    "sticky": '<path d="M5 4h14a1 1 0 0 1 1 1v9l-6 6H5a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z"/><path d="M14 20v-5a1 1 0 0 1 1-1h5"/>',
    "checklist": '<rect x="3.5" y="4.5" width="5.5" height="5.5" rx="1.4"/><path d="M4.6 17l1.6 1.6 2.9-3.1"/><path d="M12.5 7.2H20M12.5 17H20"/>',
    "clock": '<circle cx="12" cy="13" r="7.5"/><path d="M12 9.5V13l2.6 1.8"/><path d="M9.5 2.8h5"/>',
    "trash": '<path d="M5 7h14M10 7V4.5h4V7M7 7l1 13h8l1-13"/>',
    "drop": '<path d="M12 3.5c3.6 4.2 6 7.1 6 10.1a6 6 0 0 1-12 0c0-3 2.4-5.9 6-10.1z"/>',
    "pencil": '<path d="M4 20h4L19 9l-4-4L4 16z"/><path d="M13.5 6.5l4 4"/>',
    "fontplus": '<path d="M4 19l5-14h1l5 14M6 14h7"/><path d="M19 8v6M16 11h6"/>',
    "home": '<path d="M4 11l8-7 8 7v8a1 1 0 0 1-1 1h-4.5v-6h-5v6H5a1 1 0 0 1-1-1z"/>',
}

_PM_CACHE = {}


def cached_svg(name, color, size):
    key = (name, color, size)
    if key not in _PM_CACHE:
        _PM_CACHE[key] = svg_pixmap(SVG[name], color, size=size, dpr=2.0)
    return _PM_CACHE[key]


def svg_pixmap(body, color, size=20, dpr=2.0, stroke=1.8):
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" '
           f'stroke-width="{stroke}" stroke-linecap="round" stroke-linejoin="round">{body.replace("{c}", color)}</svg>')
    renderer = QSvgRenderer(QByteArray(svg.encode()))
    pm = QPixmap(int(size * dpr), int(size * dpr))
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    renderer.render(p)
    p.end()
    pm.setDevicePixelRatio(dpr)
    return pm


def glyph_pixmap(text, color, size=20, dpr=2.0, bold=False, italic=False, underline=False,
                 strike=False, bar=None):
    pm = QPixmap(int(size * dpr), int(size * dpr))
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setRenderHint(QPainter.RenderHint.TextAntialiasing)
    f = QFont("Segoe UI")
    f.setPixelSize(int(15 * dpr))
    f.setBold(bold)
    f.setItalic(italic)
    f.setUnderline(underline)
    f.setStrikeOut(strike)
    p.setFont(f)
    p.setPen(QColor(color))
    h = size * dpr
    if bar:
        p.drawText(QRectF(0, -2 * dpr, h, h - 3 * dpr), Qt.AlignmentFlag.AlignCenter, text)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(bar))
        p.drawRoundedRect(QRectF(3 * dpr, h - 5 * dpr, h - 6 * dpr, 3.2 * dpr), 1.6 * dpr, 1.6 * dpr)
    else:
        p.drawText(QRectF(0, 0, h, h), Qt.AlignmentFlag.AlignCenter, text)
    p.end()
    pm.setDevicePixelRatio(dpr)
    return pm


def make_icon(spec, th):
    """spec: ('svg', name) или ('glyph', text, {opts})."""
    icon = QIcon()
    for color, mode, state in ((th["text"], QIcon.Mode.Normal, QIcon.State.Off),
                               (th["accent"], QIcon.Mode.Normal, QIcon.State.On),
                               (th["muted"], QIcon.Mode.Disabled, QIcon.State.Off)):
        if spec[0] == "svg":
            pm = svg_pixmap(SVG[spec[1]], color)
        else:
            pm = glyph_pixmap(spec[1], color, **spec[2])
        icon.addPixmap(pm, mode, state)
    return icon


def app_icon():
    p = res("icon.ico")
    if p.exists():
        return QIcon(str(p))
    return QIcon()


# ───────────────────────────── Кодировки ─────────────────────────────
def decode_text(data: bytes):
    """Возвращает (текст, кодировка, перенос_строк)."""
    newline = "\r\n" if b"\r\n" in data else "\n"
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", "replace"), "utf-8-sig", newline
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        return data.decode("utf-16", "replace"), "utf-16", newline
    try:
        return data.decode("utf-8"), "utf-8", newline
    except UnicodeDecodeError:
        pass
    # Кириллица (cp1251) идёт «пачками» старших байтов, итальянские буквы с акцентом — по одной.
    high = [i for i, b in enumerate(data) if b >= 0x80]
    paired = sum(1 for i in high if (i + 1 < len(data) and data[i + 1] >= 0x80) or (i > 0 and data[i - 1] >= 0x80))
    enc = "cp1251" if high and paired / len(high) > 0.5 else "cp1252"
    return data.decode(enc, "replace"), enc, newline


# ───────────────────────────── DOCX: чтение ─────────────────────────────
def _num_kind(docx_doc, num_id, ilvl):
    """bullet или number для нумерации Word."""
    try:
        from docx.oxml.ns import qn
        numbering = docx_doc.part.numbering_part.element
        for num in numbering.findall(qn("w:num")):
            if num.get(qn("w:numId")) == str(num_id):
                abs_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
                for an in numbering.findall(qn("w:abstractNum")):
                    if an.get(qn("w:abstractNumId")) == abs_id:
                        for lvl in an.findall(qn("w:lvl")):
                            if lvl.get(qn("w:ilvl")) == str(ilvl):
                                fmt = lvl.find(qn("w:numFmt"))
                                if fmt is not None and fmt.get(qn("w:val")) == "bullet":
                                    return "bullet"
                                return "number"
    except Exception:
        pass
    return "bullet"


UL_TO_DOCX = {"SingleUnderline": "SINGLE", "WaveUnderline": "WAVY", "DashUnderline": "DASH",
              "DotLine": "DOTTED", "DashDotLine": "DOT_DASH", "DashDotDotLine": "DOT_DOT_DASH"}
UL_FROM_DOCX = {"WAVY": "WaveUnderline", "WAVY_HEAVY": "WaveUnderline", "WAVY_DOUBLE": "WaveUnderline",
                "DASH": "DashUnderline", "DASH_HEAVY": "DashUnderline", "DASH_LONG": "DashUnderline",
                "DOTTED": "DotLine", "DOTTED_HEAVY": "DotLine", "DOT_DASH": "DashDotLine",
                "DOT_DOT_DASH": "DashDotDotLine"}
CHECK_CHARS = {"☐": "check", "☑": "checked", "☒": "checked"}

HIGHLIGHT_MAP = {
    "YELLOW": "#FFF176", "BRIGHT_GREEN": "#A5F3A0", "TURQUOISE": "#80DEEA", "PINK": "#F8BBD0",
    "BLUE": "#90CAF9", "RED": "#EF9A9A", "DARK_BLUE": "#7986CB", "TEAL": "#80CBC4",
    "GREEN": "#81C784", "VIOLET": "#CE93D8", "DARK_RED": "#E57373", "DARK_YELLOW": "#FFD54F",
    "GRAY_25": "#E0E0E0", "GRAY_50": "#BDBDBD",
}


def load_docx(path, qdoc: QTextDocument):
    from docx import Document as DocxDocument
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    from docx.text.run import Run
    from docx.text.hyperlink import Hyperlink
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    d = DocxDocument(path)
    qdoc.clear()
    cur = QTextCursor(qdoc)
    state = {"fresh": True, "img": 0}
    active_lists = {}

    align_map = {WD_ALIGN_PARAGRAPH.CENTER: Qt.AlignmentFlag.AlignHCenter,
                 WD_ALIGN_PARAGRAPH.RIGHT: Qt.AlignmentFlag.AlignRight,
                 WD_ALIGN_PARAGRAPH.JUSTIFY: Qt.AlignmentFlag.AlignJustify}

    def new_block(bf, cf):
        if state["fresh"]:
            cur.setBlockFormat(bf)
            cur.setBlockCharFormat(cf)
            cur.setCharFormat(cf)
            state["fresh"] = False
        else:
            cur.insertBlock(bf, cf)

    def run_format(run, base):
        cf = QTextCharFormat(base)
        if run.bold is True:
            cf.setFontWeight(700)
        elif run.bold is False:
            cf.setFontWeight(400)
        if run.italic:
            cf.setFontItalic(True)
        if run.underline:
            name = getattr(run.underline, "name", "SINGLE") if run.underline is not True else "SINGLE"
            st = UL_FROM_DOCX.get(name, "SingleUnderline")
            cf.setUnderlineStyle(getattr(QTextCharFormat.UnderlineStyle, st))
            try:
                u = run._element.rPr.find(qn("w:u"))
                col = u.get(qn("w:color")) if u is not None else None
                if col and col.lower() != "auto":
                    cf.setUnderlineColor(QColor("#" + col))
            except Exception:
                pass
        f = run.font
        if f.strike:
            cf.setFontStrikeOut(True)
        if f.size:
            cf.setFontPointSize(f.size.pt)
        if f.name:
            cf.setFontFamilies([f.name])
        try:
            if f.color is not None and f.color.type is not None and f.color.rgb is not None:
                rgb = str(f.color.rgb)
                if rgb not in ("000000",):
                    cf.setForeground(QColor("#" + rgb))
        except Exception:
            pass
        try:
            if f.highlight_color is not None:
                name = str(f.highlight_color).split(".")[-1].split(" ")[0]
                if name in HIGHLIGHT_MAP:
                    cf.setBackground(QColor(HIGHLIGHT_MAP[name]))
            rpr = run._element.rPr
            if rpr is not None:
                shd = rpr.find(qn("w:shd"))
                if shd is not None:
                    fill = shd.get(qn("w:fill"))
                    if fill and fill.lower() not in ("auto", "ffffff"):
                        cf.setBackground(QColor("#" + fill))
        except Exception:
            pass
        return cf

    def insert_run(run, base, href=None):
        cf = run_format(run, base)
        if href:
            cf.setAnchor(True)
            cf.setAnchorHref(href)
            cf.setFontUnderline(True)
            cf.setForeground(QColor(LINK_COLOR))
        # картинки внутри run
        for blip in run._element.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            try:
                blob = d.part.related_parts[rid].blob
                img = QImage.fromData(blob)
                if img.isNull():
                    continue
                w, h = img.width(), img.height()
                parent = blip
                ext = None
                while parent is not None:
                    ext = parent.find(qn("wp:extent"))
                    if ext is not None:
                        break
                    parent = parent.getparent()
                if ext is not None:
                    w = int(ext.get("cx")) / 9525
                    h = int(ext.get("cy")) / 9525
                if w > 640:
                    h = h * 640 / w
                    w = 640
                state["img"] += 1
                name = f"docx-img-{state['img']}"
                qdoc.addResource(QTextDocument.ResourceType.ImageResource, QUrl(name), img)
                imf = QTextImageFormat()
                imf.setName(name)
                imf.setWidth(w)
                imf.setHeight(h)
                cur.insertImage(imf)
            except Exception:
                continue
        text = run.text
        if text:
            cur.insertText(text.replace("\n", " "), cf)

    def add_paragraph(para):
        bf = QTextBlockFormat()
        cf = QTextCharFormat()
        style = (para.style.name if para.style is not None else "") or ""
        level = 0
        m = re.match(r"(?i)heading\s*(\d)", style)
        if m:
            level = min(int(m.group(1)), 3)
        elif style.lower() == "title":
            level = 1
        elif style.lower() == "subtitle":
            level = 2
        if level:
            bf.setHeadingLevel(level)
            cf.setFontPointSize(HEADING_SIZES[level])
            cf.setFontWeight(700)
        if para.alignment in align_map:
            bf.setAlignment(align_map[para.alignment])

        # списки
        kind, ilvl = None, 0
        ppr = para._p.pPr
        numpr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if numpr is not None and numpr.find(qn("w:numId")) is not None:
            num_id = numpr.find(qn("w:numId")).get(qn("w:val"))
            il = numpr.find(qn("w:ilvl"))
            ilvl = int(il.get(qn("w:val"))) if il is not None else 0
            if num_id != "0":
                kind = _num_kind(d, num_id, ilvl)
        elif style.lower().startswith("list bullet"):
            kind = "bullet"
            m2 = re.search(r"(\d)$", style)
            ilvl = int(m2.group(1)) - 1 if m2 else 0
        elif style.lower().startswith("list number"):
            kind = "number"
            m2 = re.search(r"(\d)$", style)
            ilvl = int(m2.group(1)) - 1 if m2 else 0

        mcheck = re.match(r"^\s*([☐☑☒])\s?", para.text or "")
        if mcheck and not kind:
            kind, ilvl = "check", 0
        new_block(bf, cf)
        if kind == "check":
            for k in [k for k in active_lists if k[0] != "check"]:
                del active_lists[k]
        if kind:
            for k in [k for k in active_lists if k[1] > ilvl]:
                del active_lists[k]
            key = (kind, ilvl)
            if key in active_lists:
                active_lists[key].add(cur.block())
            else:
                lf = QTextListFormat()
                lf.setIndent(ilvl + 1)
                if kind in ("bullet", "check"):
                    lf.setStyle([QTextListFormat.Style.ListDisc, QTextListFormat.Style.ListCircle,
                                 QTextListFormat.Style.ListSquare][ilvl % 3])
                else:
                    lf.setStyle(QTextListFormat.Style.ListDecimal)
                active_lists[key] = cur.createList(lf)
            if kind == "check":
                mbf = QTextBlockFormat()
                mbf.setMarker(CHECKED if mcheck and CHECK_CHARS[mcheck.group(1)] == "checked" else UNCHECKED)
                cur.mergeBlockFormat(mbf)
        else:
            active_lists.clear()

        for item in para.iter_inner_content():
            if isinstance(item, Run):
                insert_run(item, cf)
            elif isinstance(item, Hyperlink):
                href = None
                try:
                    href = item.url or None
                except Exception:
                    pass
                for r in item.runs:
                    insert_run(r, cf, href)
        if mcheck:  # убираем «☐ » из текста — галочка теперь настоящая
            c = QTextCursor(cur.block())
            c.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor, len(mcheck.group(0)))
            c.removeSelectedText()

    def add_table(table):
        active_lists.clear()
        rows = len(table.rows)
        cols = max((len(r.cells) for r in table.rows), default=0)
        if not rows or not cols:
            return
        tf = QTextTableFormat()
        tf.setBorder(1)
        tf.setBorderStyle(QTextFrameFormat.BorderStyle.BorderStyle_Solid)
        tf.setBorderBrush(QBrush(QColor("#C7CCD6")))
        tf.setBorderCollapse(True)
        tf.setCellPadding(6)
        tf.setCellSpacing(0)
        tf.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        qt = cur.insertTable(rows, cols, tf)
        for r, row in enumerate(table.rows):
            seen = set()
            for c, cell in enumerate(row.cells):
                if c >= cols:
                    break
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                cc = qt.cellAt(r, c).firstCursorPosition()
                for i, p in enumerate(cell.paragraphs):
                    if i:
                        cc.insertBlock()
                    bold = bool(p.runs) and all(run.bold for run in p.runs if run.text.strip())
                    cf = QTextCharFormat()
                    if bold:
                        cf.setFontWeight(700)
                    cc.insertText(p.text, cf)
        cur.movePosition(QTextCursor.MoveOperation.End)
        state["fresh"] = True

    body = d.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            add_paragraph(Paragraph(child, d))
        elif child.tag == qn("w:tbl"):
            add_table(Table(child, d))


# ───────────────────────────── DOCX: запись ─────────────────────────────
def _image_bytes(qdoc, name):
    data = qdoc.resource(QTextDocument.ResourceType.ImageResource, QUrl(name))
    if data is None:
        return None
    if isinstance(data, QPixmap):
        data = data.toImage()
    if isinstance(data, (QByteArray, bytes)):
        img = QImage.fromData(bytes(data))
    elif isinstance(data, QImage):
        img = data
    else:
        return None
    if img.isNull():
        return None
    buf = QBuffer()
    buf.open(QIODevice.OpenModeFlag.WriteOnly)
    img.save(buf, "PNG")
    return bytes(buf.data()), img.width(), img.height()


def save_docx(qdoc: QTextDocument, path):
    from docx import Document as DocxDocument
    from docx.shared import Pt, RGBColor, Emu
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.enum.text import WD_UNDERLINE

    d = DocxDocument()
    normal = d.styles["Normal"]
    normal.font.name = DEFAULT_FONT
    normal.font.size = Pt(DEFAULT_SIZE)

    def write_block(block, para):
        bf = block.blockFormat()
        lvl = bf.headingLevel()
        lst = block.textList()
        marker = bf.marker()
        if lvl:
            para.style = d.styles[f"Heading {min(lvl, 3)}"]
        elif lst is not None and marker != NO_MARKER:
            para.add_run("☑ " if marker == CHECKED else "☐ ")
            para.paragraph_format.left_indent = Pt(18 * max(0, lst.format().indent() - 1))
        elif lst is not None:
            lf = lst.format()
            depth = max(1, min(lf.indent(), 3))
            bullet = lf.style() in (QTextListFormat.Style.ListDisc, QTextListFormat.Style.ListCircle,
                                    QTextListFormat.Style.ListSquare)
            base = "List Bullet" if bullet else "List Number"
            para.style = d.styles[base if depth == 1 else f"{base} {depth}"]
        al = bf.alignment()
        if al & Qt.AlignmentFlag.AlignHCenter:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif al & Qt.AlignmentFlag.AlignJustify:
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif al & Qt.AlignmentFlag.AlignRight:
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        it = block.begin()
        while not it.atEnd():
            frag = it.fragment()
            it += 1
            if not frag.isValid():
                continue
            cf = frag.charFormat()
            text = frag.text()
            if cf.isImageFormat():
                imf = cf.toImageFormat()
                got = _image_bytes(qdoc, imf.name())
                if got:
                    data, iw, ih = got
                    w = imf.width() or iw
                    for _ in range(len(text)):
                        para.add_run().add_picture(io.BytesIO(data), width=Emu(int(w * 9525)))
                continue
            run = para.add_run(text.replace(" ", "\n"))
            if cf.fontWeight() >= 600:
                run.bold = True
            if cf.fontItalic():
                run.italic = True
            ul = cf.underlineStyle()
            if ul != QTextCharFormat.UnderlineStyle.NoUnderline and not cf.isAnchor():
                run.underline = getattr(WD_UNDERLINE, UL_TO_DOCX.get(ul.name, "SINGLE"))
                if cf.underlineColor().isValid() and cf.hasProperty(QTextFormat.Property.TextUnderlineColor):
                    uc = cf.underlineColor()
                    u = run._element.get_or_add_rPr().find(qn("w:u"))
                    if u is not None:
                        u.set(qn("w:color"), f"{uc.red():02X}{uc.green():02X}{uc.blue():02X}")
            if cf.fontStrikeOut():
                run.font.strike = True
            size = cf.fontPointSize()
            if size > 0 and (lvl == 0 and abs(size - DEFAULT_SIZE) > 0.1):
                run.font.size = Pt(size)
            fams = cf.fontFamilies() or []
            fam = fams[0] if fams else ""
            if fam and fam != DEFAULT_FONT:
                run.font.name = fam
            if cf.isAnchor():
                run.underline = True
                run.font.color.rgb = RGBColor.from_string(LINK_COLOR[1:])
            elif cf.foreground().style() != Qt.BrushStyle.NoBrush:
                c = cf.foreground().color()
                run.font.color.rgb = RGBColor(c.red(), c.green(), c.blue())
            if cf.background().style() != Qt.BrushStyle.NoBrush:
                c = cf.background().color()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), f"{c.red():02X}{c.green():02X}{c.blue():02X}")
                run._element.get_or_add_rPr().append(shd)
            href = cf.anchorHref() if cf.isAnchor() else ""
            if href:
                # настоящая гиперссылка Word: текст виден, адрес спрятан
                try:
                    rid = para.part.relate_to(href, RT.HYPERLINK, is_external=True)
                    link = OxmlElement("w:hyperlink")
                    link.set(qn("r:id"), rid)
                    para._p.remove(run._r)
                    link.append(run._r)
                    para._p.append(link)
                except Exception:
                    pass

    def write_table(table):
        rows, cols = table.rows(), table.columns()
        t = d.add_table(rows=rows, cols=cols)
        t.style = d.styles["Table Grid"]
        for r in range(rows):
            for c in range(cols):
                cell = table.cellAt(r, c)
                cc = QTextCursor(qdoc)
                cc.setPosition(cell.firstCursorPosition().position())
                cc.setPosition(cell.lastCursorPosition().position(), QTextCursor.MoveMode.KeepAnchor)
                parts = cc.selectedText().split(" ")
                dc = t.cell(r, c)
                dc.paragraphs[0].text = parts[0].replace(" ", "\n")
                for extra in parts[1:]:
                    dc.add_paragraph(extra.replace(" ", "\n"))

    block = qdoc.begin()
    while block.isValid():
        table = QTextCursor(block).currentTable()
        if table is not None:
            write_table(table)
            nxt = qdoc.findBlock(table.lastPosition() + 1)
            if not nxt.isValid() or nxt.position() <= block.position():
                break
            block = nxt
            continue
        write_block(block, d.add_paragraph())
        block = block.next()

    # убираем пустой первый абзац шаблона, если он есть
    d.save(path)


# ───────────────────────────── Экспорт ─────────────────────────────
EXPORT_FORMATS = [
    ("docx", "Word (.docx)", "Для Word и Google Документов"),
    ("odt", "OpenDocument (.odt)", "LibreOffice, Google Документы"),
    ("pdf", "PDF (.pdf)", "Отправить или распечатать"),
    ("rtf", "RTF (.rtf)", "Открывается почти везде"),
    ("html", "Веб-страница (.html)", "С картинками внутри"),
    ("md", "Markdown (.md)", "Notion, Obsidian, GitHub"),
    ("txt", "Простой текст (.txt)", "Без оформления"),
]


def iter_body(qdoc):
    """По очереди: ('block', QTextBlock) или ('table', QTextTable)."""
    block = qdoc.begin()
    while block.isValid():
        table = QTextCursor(block).currentTable()
        if table is not None:
            yield "table", table
            nxt = qdoc.findBlock(table.lastPosition() + 1)
            if not nxt.isValid() or nxt.position() <= block.position():
                break
            block = nxt
            continue
        yield "block", block
        block = block.next()


def cell_blocks(qdoc, cell):
    b = cell.firstCursorPosition().block()
    end = cell.lastCursorPosition().position()
    while b.isValid() and b.position() <= end:
        yield b
        b = b.next()


def block_fragments(block):
    it = block.begin()
    while not it.atEnd():
        f = it.fragment()
        it += 1
        if f.isValid():
            yield f


def print_document(qdoc, printer):
    printer.setPageLayout(QPageLayout(QPageSize(QPageSize.PageSizeId.A4), QPageLayout.Orientation.Portrait,
                                      QMarginsF(18, 18, 18, 18), QPageLayout.Unit.Millimeter))
    clone = qdoc.clone()
    clone.setDocumentMargin(0)
    clone.print_(printer)


def _rtf_text(text):
    out = []
    for ch in text:
        o = ord(ch)
        if ch in "\\{}":
            out.append("\\" + ch)
        elif ch == "\u2028":
            out.append("\\line ")
        elif ch == "\t":
            out.append("\\tab ")
        elif ch == "\ufffc":
            continue
        elif o < 128:
            out.append(ch)
        else:
            units = [o] if o <= 0xFFFF else [0xD800 + ((o - 0x10000) >> 10), 0xDC00 + ((o - 0x10000) & 0x3FF)]
            for u in units:
                out.append(f"\\u{u - 65536 if u > 32767 else u}?")
    return "".join(out)


def save_rtf(qdoc, path):
    fonts = [DEFAULT_FONT]
    colors = []

    def fi(name):
        if name not in fonts:
            fonts.append(name)
        return fonts.index(name)

    def ci(c):
        key = (c.red(), c.green(), c.blue())
        if key not in colors:
            colors.append(key)
        return colors.index(key) + 1

    def frag_rtf(frag, base_size, heading):
        cf = frag.charFormat()
        if cf.isImageFormat():
            imf = cf.toImageFormat()
            got = _image_bytes(qdoc, imf.name())
            if not got:
                return ""
            data, iw, ih = got
            w = imf.width() or iw
            h = imf.height() or ih
            return (f"{{\\pict\\pngblip\\picw{iw}\\pich{ih}\\picwgoal{round(w * 15)}\\pichgoal{round(h * 15)} "
                    f"{data.hex()}}}")
        codes = []
        fams = cf.fontFamilies() or []
        if fams:
            codes.append(f"\\f{fi(fams[0])}")
        size = cf.fontPointSize() or base_size
        codes.append(f"\\fs{round(size * 2)}")
        if cf.fontWeight() >= 600 or heading:
            codes.append("\\b")
        if cf.fontItalic():
            codes.append("\\i")
        ul = cf.underlineStyle()
        if cf.isAnchor():
            codes.append("\\ul")
        elif ul != QTextCharFormat.UnderlineStyle.NoUnderline:
            codes.append({"WaveUnderline": "\\ulwave", "DashUnderline": "\\uldash", "DotLine": "\\uld",
                          "DashDotLine": "\\uldashd", "DashDotDotLine": "\\uldashdd"}.get(ul.name, "\\ul"))
            if cf.hasProperty(QTextFormat.Property.TextUnderlineColor):
                codes.append(f"\\ulc{ci(cf.underlineColor())}")
        if cf.fontStrikeOut():
            codes.append("\\strike")
        if cf.isAnchor():
            codes.append(f"\\cf{ci(QColor(LINK_COLOR))}")
        elif cf.foreground().style() != Qt.BrushStyle.NoBrush:
            codes.append(f"\\cf{ci(cf.foreground().color())}")
        if cf.background().style() != Qt.BrushStyle.NoBrush:
            n = ci(cf.background().color())
            codes.append(f"\\chcbpat{n}\\cb{n}")
        inner = "{" + "".join(codes) + " " + _rtf_text(frag.text()) + "}"
        href = cf.anchorHref() if cf.isAnchor() else ""
        if href:
            inner = '{\\field{\\*\\fldinst HYPERLINK "' + href.replace('"', "%22") + '"}{\\fldrslt ' + inner + "}}"
        return inner

    def para(block, in_table=False):
        bf = block.blockFormat()
        parts = ["\\pard\\plain"]
        if in_table:
            parts.append("\\intbl")
        al = bf.alignment()
        if al & Qt.AlignmentFlag.AlignHCenter:
            parts.append("\\qc")
        elif al & Qt.AlignmentFlag.AlignJustify:
            parts.append("\\qj")
        elif al & Qt.AlignmentFlag.AlignRight:
            parts.append("\\qr")
        lvl = min(bf.headingLevel(), 3)
        base = HEADING_SIZES[lvl] if lvl else DEFAULT_SIZE
        parts.append("\\sb240\\sa120" if lvl else "\\sa120")
        prefix = ""
        lst = block.textList()
        if lst is not None:
            lf = lst.format()
            parts.append(f"\\li{360 * lf.indent()}\\fi-360")
            bullet = lf.style() in (QTextListFormat.Style.ListDisc, QTextListFormat.Style.ListCircle,
                                    QTextListFormat.Style.ListSquare)
            mk = bf.marker()
            if mk != NO_MARKER:
                bullet = True
            box = "\\u9745?" if mk == CHECKED else "\\u9744?"
            prefix = (f"{{\\fs{round(base * 2)} {box}}}" if mk != NO_MARKER else
                      f"{{\\fs{round(base * 2)} \\u8226?}}" if bullet else
                      f"{{\\fs{round(base * 2)} {_rtf_text(lst.itemText(block))}}}") + "\\tab "
        body = "".join(frag_rtf(f, base, bool(lvl)) for f in block_fragments(block))
        return "".join(parts) + " " + prefix + body

    out = []
    for kind, obj in iter_body(qdoc):
        if kind == "block":
            out.append(para(obj) + "\\par\n")
        else:
            rows, cols = obj.rows(), obj.columns()
            width = 9000 // max(cols, 1)
            for r in range(rows):
                row = ["\\trowd\\trgaph108"]
                for c in range(cols):
                    row.append("\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10"
                               f"\\clbrdrr\\brdrs\\brdrw10\\cellx{width * (c + 1)}")
                for c in range(cols):
                    blocks = list(cell_blocks(qdoc, obj.cellAt(r, c)))
                    row.append("\\par ".join(para(b, True) for b in blocks) + "\\cell ")
                row.append("\\row\n")
                out.append("".join(row))
    head = ["{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1{\\fonttbl"]
    for i, name in enumerate(fonts):
        head.append(f"{{\\f{i}\\fnil\\fcharset0 {name};}}")
    head.append("}{\\colortbl;")
    for r, g, b in colors:
        head.append(f"\\red{r}\\green{g}\\blue{b};")
    head.append("}\n\\viewkind4\n")
    Path(path).write_text("".join(head) + "".join(out) + "}", encoding="ascii")


def save_html(qdoc, path):
    import base64
    text = qdoc.toHtml()
    for name in set(re.findall(r'<img[^>]*src="([^"]+)"', text)):
        got = _image_bytes(qdoc, html.unescape(name))
        if got:
            text = text.replace(f'src="{name}"', 'src="data:image/png;base64,' + base64.b64encode(got[0]).decode() + '"')
    Path(path).write_text(text, encoding="utf-8")


def export_document(qdoc, path, fmt):
    if fmt == "docx":
        save_docx(qdoc, path)
    elif fmt == "odt":
        from PySide6.QtGui import QTextDocumentWriter
        # Qt путает разметку, если таблица идёт сразу после списка — разделяем пустым абзацем
        doc = qdoc.clone()
        for kind, obj in list(iter_body(doc)):
            if kind == "table" and obj.firstPosition() > 1:
                prev = doc.findBlock(obj.firstPosition() - 1)
                if prev.textList() is not None:
                    c = QTextCursor(prev)
                    c.movePosition(QTextCursor.MoveOperation.EndOfBlock)
                    c.insertBlock(QTextBlockFormat(), QTextCharFormat())
        writer = QTextDocumentWriter(path, QByteArray(b"odf"))
        if not writer.write(doc):
            raise OSError("не получилось записать файл")
    elif fmt == "pdf":
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName(path)
        print_document(qdoc, printer)
    elif fmt == "rtf":
        save_rtf(qdoc, path)
    elif fmt == "html":
        save_html(qdoc, path)
    elif fmt == "md":
        Path(path).write_text(qdoc.toMarkdown(), encoding="utf-8")
    else:
        Path(path).write_bytes(qdoc.toPlainText().replace("\u2028", "\n").replace("\n", "\r\n").encode("utf-8"))


def load_document(path):
    """Открывает файл в отдельный QTextDocument (для экспорта прямо из проекта)."""
    p = Path(path)
    d = QTextDocument()
    f = QFont(DEFAULT_FONT)
    f.setPointSizeF(DEFAULT_SIZE)
    d.setDefaultFont(f)
    ext = p.suffix.lower()
    if ext == ".docx":
        load_docx(str(p), d)
    else:
        text, _, _ = decode_text(p.read_bytes())
        if ext in MD_EXTS:
            d.setMarkdown(text)
        elif ext in HTML_EXTS:
            d.setHtml(text)
        else:
            d.setPlainText(text)
    return d


def export_menu(parent, callback, th):
    m = styled_menu(parent)
    m.setTitle("Экспорт")
    m.setIcon(make_icon(("svg", "export"), th))
    for fmt, label, hint in EXPORT_FORMATS:
        a = m.addAction(f"{label}   ·   {hint}")
        a.triggered.connect(lambda _=False, f=fmt: callback(f))
    return m


# ───────────────────────────── Виджеты ─────────────────────────────
class RoundButton(QToolButton):
    def __init__(self, tip, checkable=False, size=36):
        super().__init__()
        self.setToolTip(tip)
        self.setCheckable(checkable)
        self.setFixedSize(size, size)
        self.setIconSize(QSize(20, 20))
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAutoRaise(True)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.spec = None


def styled_menu(parent=None):
    m = QMenu(parent)
    m.setWindowFlags(m.windowFlags() | Qt.WindowType.FramelessWindowHint | Qt.WindowType.NoDropShadowWindowHint)
    m.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
    return m


def recent_colors():
    v = QSettings(APP_NAME, APP_NAME).value("recent_colors", [])
    if isinstance(v, str):
        v = [v]
    return [c for c in (v or []) if QColor(c).isValid()][:12]


def remember_color(c):
    name = QColor(c).name(QColor.NameFormat.HexRgb).upper()
    lst = [x for x in recent_colors() if x.upper() != name]
    QSettings(APP_NAME, APP_NAME).setValue("recent_colors", ([name] + lst)[:12])


def pick_custom_color(parent, initial="#3D5AFE", title="Свой цвет"):
    dlg = QColorDialog(QColor(initial), parent)
    dlg.setWindowTitle(title)
    for i, c in enumerate(recent_colors()[:16]):
        QColorDialog.setCustomColor(i, QColor(c))
    if dlg.exec() == QDialog.DialogCode.Accepted and dlg.selectedColor().isValid():
        c = dlg.selectedColor()
        remember_color(c)
        return c
    return None


class ColorRow(QWidget):
    """Ряд кружков-цветов: готовые + свои (недавние) + «+» для любого цвета."""
    picked = Signal(object)

    def __init__(self, colors, current=None, allow_none=False, none_label="авто", per_row=8, size=26):
        super().__init__()
        self.current = (current or "").upper()
        grid = QGridLayout(self)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(7)
        items = []
        if allow_none:
            items.append(("", none_label))
        items += [(c, c) for c in colors]
        rec = [c for c in recent_colors() if c.upper() not in [x.upper() for x in colors]]
        items += [(c, c) for c in rec[:per_row]]
        if self.current and self.current not in [i[0].upper() for i in items]:
            items.append((self.current, self.current))
        for i, (col, tip) in enumerate(items):
            b = QToolButton()
            b.setFixedSize(size if col else size + 14, size)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setToolTip(tip)
            sel = col.upper() == self.current
            ring = "#3D5AFE" if sel else "rgba(127,127,127,0.35)"
            if col:
                b.setStyleSheet(f"QToolButton{{background:{col};border-radius:{size // 2}px;"
                                f"border:{3 if sel else 1}px solid {ring};}}"
                                f"QToolButton:hover{{border:2px solid #3D5AFE;}}")
            else:
                b.setText(none_label)
                b.setStyleSheet(f"QToolButton{{background:transparent;border-radius:{size // 2}px;font-size:7.5pt;"
                                f"border:{3 if sel else 1}px solid {ring};}}")
            b.clicked.connect(lambda _=False, c=col: self.picked.emit(QColor(c) if c else None))
            grid.addWidget(b, i // per_row, i % per_row)
        plus = QToolButton()
        plus.setFixedSize(size, size)
        plus.setCursor(Qt.CursorShape.PointingHandCursor)
        plus.setToolTip("Свой цвет — любой")
        plus.setText("+")
        plus.setStyleSheet(f"QToolButton{{border:1px dashed rgba(127,127,127,0.8);border-radius:{size // 2}px;"
                           f"font-size:12pt;}}")
        plus.clicked.connect(self._custom)
        n = len(items)
        grid.addWidget(plus, n // per_row, n % per_row)

    def _custom(self):
        c = pick_custom_color(self.window(), self.current or "#3D5AFE")
        if c is not None:
            self.picked.emit(c)


class SwatchMenu(QMenu):
    """Палитра цветов в стиле Material (с «своими» цветами)."""
    picked = Signal(object)  # QColor или None (сбросить)

    COLORS = ["#1A1D24", "#5F6368", "#E5484D", "#F76B15", "#F5B400", "#30A46C",
              "#12A594", "#0090FF", "#3D5AFE", "#8E4EC6", "#D6409F", "#8D6E63"]
    HIGHLIGHTS = ["#FFF176", "#FFE0B2", "#FFCDD2", "#F8BBD0", "#E1BEE7", "#C5CAE9",
                  "#BBDEFB", "#B2EBF2", "#C8E6C9", "#DCEDC8", "#F0F4C3", "#E0E0E0"]

    def __init__(self, parent, highlight=False, colors=None, reset=True):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.FramelessWindowHint | Qt.WindowType.NoDropShadowWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.colors = colors or (self.HIGHLIGHTS if highlight else self.COLORS)
        self.highlight = highlight
        self.reset = reset
        self.aboutToShow.connect(self._build)  # перестраиваем, чтобы видеть свежие «свои» цвета

    def _build(self):
        self.clear()
        box = QWidget()
        box.setObjectName("SwatchBox")
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 10, 10, 10)
        v.setSpacing(8)
        row = ColorRow(self.colors, per_row=6)
        row.picked.connect(self._pick)
        v.addWidget(row)
        if self.reset:
            rb = QPushButton("Без выделения" if self.highlight else "Цвет по умолчанию")
            rb.setObjectName("SwatchReset")
            rb.setCursor(Qt.CursorShape.PointingHandCursor)
            rb.clicked.connect(lambda: self._pick(None))
            v.addWidget(rb)
        act = QWidgetAction(self)
        act.setDefaultWidget(box)
        self.addAction(act)

    def _pick(self, c):
        self.picked.emit(c)
        self.close()


IMG_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
NO_MARKER = QTextBlockFormat.MarkerType.NoMarker
UNCHECKED = QTextBlockFormat.MarkerType.Unchecked
CHECKED = QTextBlockFormat.MarkerType.Checked
BULLET_STYLES = (QTextListFormat.Style.ListDisc, QTextListFormat.Style.ListCircle, QTextListFormat.Style.ListSquare)


def blocks_in(cur):
    doc = cur.document()
    b = doc.findBlock(cur.selectionStart())
    last = doc.findBlock(cur.selectionEnd())
    while b.isValid():
        yield b
        if b == last:
            break
        b = b.next()


def make_list(cur, kind, indent=1):
    """kind: bullet | number | check | checked"""
    lf = QTextListFormat()
    lf.setIndent(indent)
    lf.setStyle(QTextListFormat.Style.ListDecimal if kind == "number" else QTextListFormat.Style.ListDisc)
    cur.createList(lf)
    if kind in ("check", "checked"):
        for b in list(blocks_in(cur)):
            bf = QTextBlockFormat()
            bf.setMarker(CHECKED if kind == "checked" else UNCHECKED)
            QTextCursor(b).mergeBlockFormat(bf)


def unlist_block(block):
    lst = block.textList()
    if lst is not None:
        lst.remove(block)
    bf = block.blockFormat()
    bf.setIndent(0)
    bf.clearProperty(QTextFormat.Property.BlockMarker)
    QTextCursor(block).setBlockFormat(bf)


def toggle_check(block):
    checked = block.blockFormat().marker() == CHECKED
    c = QTextCursor(block)
    c.beginEditBlock()
    bf = QTextBlockFormat()
    bf.setMarker(UNCHECKED if checked else CHECKED)
    c.mergeBlockFormat(bf)
    c.movePosition(QTextCursor.MoveOperation.StartOfBlock)
    c.movePosition(QTextCursor.MoveOperation.EndOfBlock, QTextCursor.MoveMode.KeepAnchor)
    cf = QTextCharFormat()
    cf.setFontStrikeOut(not checked)
    c.mergeCharFormat(cf)
    c.endEditBlock()


AUTO_LISTS = {"1.": "number", "1)": "number", "-": "bullet", "*": "bullet", "•": "bullet",
              "[]": "check", "[ ]": "check", "[x]": "checked", "[х]": "checked"}


class SmartEdit(QTextEdit):
    """Общие удобства для основного текста и заметок: списки как в Word, чек-листы, ссылки."""

    def focusInEvent(self, e):
        super().focusInEvent(e)
        w = self.window()
        if hasattr(w, "set_target"):
            w.set_target(self)

    def _marker_block(self, pos):
        block = self.cursorForPosition(pos).block()
        if block.textList() is None or block.blockFormat().marker() == NO_MARKER:
            return None
        r = self.cursorRect(QTextCursor(block))
        if r.left() - 30 <= pos.x() <= r.left() + 1 and r.top() - 4 <= pos.y() <= r.bottom() + 4:
            return block
        return None

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            block = self._marker_block(e.position().toPoint())
            if block is not None:
                toggle_check(block)
                return
        super().mousePressEvent(e)

    def keyPressEvent(self, e):
        cur = self.textCursor()
        key = e.key()
        mods = e.modifiers()
        enter = key in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and not (mods & Qt.KeyboardModifier.ShiftModifier)
        lst = cur.currentList()
        win = self.window()

        # Enter на пустом пункте — выходим из списка (как в Word)
        if enter and lst is not None and not cur.hasSelection() and not cur.block().text().strip():
            unlist_block(cur.block())
            return
        if enter and cur.blockFormat().headingLevel() and cur.atBlockEnd():
            cf = QTextCharFormat()
            cf.setFontWeight(400)
            cf.setFontPointSize(DEFAULT_SIZE)
            bf = QTextBlockFormat(cur.blockFormat())
            bf.setHeadingLevel(0)
            cur.insertBlock(bf, cf)
            self.setTextCursor(cur)
            return
        if enter and lst is not None and cur.blockFormat().marker() != NO_MARKER:
            super().keyPressEvent(e)
            c = self.textCursor()
            bf = QTextBlockFormat()
            bf.setMarker(UNCHECKED)
            c.mergeBlockFormat(bf)
            cf = self.currentCharFormat()
            cf.setFontStrikeOut(False)
            self.setCurrentCharFormat(cf)
            return
        # Backspace в начале пункта — убираем маркер, текст остаётся
        if key == Qt.Key.Key_Backspace and lst is not None and cur.atBlockStart() and not cur.hasSelection():
            unlist_block(cur.block())
            return
        if key in (Qt.Key.Key_Tab, Qt.Key.Key_Backtab) and lst is not None and hasattr(win, "change_list_indent"):
            win.change_list_indent(-1 if key == Qt.Key.Key_Backtab else +1)
            return
        # «1. », «- », «[] » в начале строки сами превращаются в список
        if key == Qt.Key.Key_Space and lst is None and not cur.hasSelection() \
                and not (mods & Qt.KeyboardModifier.ControlModifier):
            before = cur.block().text()[:cur.positionInBlock()]
            kind = AUTO_LISTS.get(before.lower())
            if kind:
                c = QTextCursor(cur)
                c.beginEditBlock()
                c.movePosition(QTextCursor.MoveOperation.StartOfBlock, QTextCursor.MoveMode.KeepAnchor)
                c.removeSelectedText()
                make_list(c, kind)
                c.endEditBlock()
                if hasattr(win, "sync_toolbar"):
                    win.sync_toolbar()
                return
        # текст, набранный сразу после ссылки, не должен становиться частью ссылки
        if e.text() and e.text().isprintable() and not cur.hasSelection():
            cf = self.currentCharFormat()
            if cf.isAnchor():
                after = QTextCursor(self.document())
                after.setPosition(min(cur.position() + 1, self.document().characterCount() - 1))
                if cur.atBlockEnd() or not after.charFormat().isAnchor() \
                        or after.charFormat().anchorHref() != cf.anchorHref():
                    self.setCurrentCharFormat(strip_link_format(cf))
        super().keyPressEvent(e)

    # ── ссылки и чек-боксы под мышкой ──
    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        href = self.anchorAt(pos)
        ctrl = bool(e.modifiers() & Qt.KeyboardModifier.ControlModifier)
        if (href and ctrl) or self._marker_block(pos) is not None:
            self.viewport().setCursor(Qt.CursorShape.PointingHandCursor)
        else:
            self.viewport().setCursor(Qt.CursorShape.IBeamCursor)
        if href != getattr(self, "_hover_href", ""):
            self._hover_href = href
            if href:
                QToolTip.showText(e.globalPosition().toPoint(), f"{href}\nCtrl + клик — открыть", self)
            else:
                QToolTip.hideText()
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        href = self.anchorAt(e.position().toPoint())
        if href and e.button() == Qt.MouseButton.LeftButton and \
                e.modifiers() & Qt.KeyboardModifier.ControlModifier and not self.textCursor().hasSelection():
            open_href(href, self.window())
            return
        super().mouseReleaseEvent(e)


class Editor(SmartEdit):
    fileDropped = Signal(str)

    def mousePressEvent(self, e):
        w = self.window()
        if hasattr(w, "stickers"):
            w.stickers.select(None)
        super().mousePressEvent(e)

    def canInsertFromMimeData(self, source):
        if source.hasUrls() or source.hasImage():
            return True
        return super().canInsertFromMimeData(source)

    def insertFromMimeData(self, source):
        if source.hasUrls():
            for url in source.urls():
                if url.isLocalFile():
                    p = url.toLocalFile()
                    if Path(p).suffix.lower() in IMG_EXTS:
                        self.window().insert_image_file(p)
                    else:
                        self.fileDropped.emit(p)
                    return
        if source.hasImage():
            img = QImage(source.imageData())
            if not img.isNull():
                self.window().insert_qimage(img)
                return
        super().insertFromMimeData(source)

    def dropEvent(self, e):
        md = e.mimeData()
        imgs = [u.toLocalFile() for u in md.urls() if u.isLocalFile()
                and Path(u.toLocalFile()).suffix.lower() in IMG_EXTS] if md.hasUrls() else []
        if imgs:
            pos = e.position().toPoint()
            m = styled_menu(self)
            in_text = m.addAction("Вставить в текст")
            sticker = m.addAction("Прикрепить стикером (видно только в Folio)")
            chosen = m.exec(self.viewport().mapToGlobal(pos))
            if chosen is sticker:
                for i, p in enumerate(imgs):
                    self.window().stickers.add_photo(p, QPoint(pos.x() + 24 * i, pos.y() + 24 * i))
                e.setDropAction(Qt.DropAction.CopyAction)
                e.accept()
                return
            if chosen is not in_text:
                e.ignore()
                return
        super().dropEvent(e)


def strip_link_format(cf):
    f = QTextCharFormat(cf)
    f.setAnchor(False)
    f.clearProperty(QTextFormat.Property.AnchorHref)
    f.clearProperty(QTextFormat.Property.AnchorName)
    f.clearForeground()
    f.setFontUnderline(False)
    return f


def normalize_url(text):
    text = text.strip()
    if not text:
        return ""
    if re.match(r"^[a-zA-Z]:[\\/]", text) or text.startswith("\\\\"):
        return QUrl.fromLocalFile(text).toString()
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", text):
        return text
    return QUrl.fromUserInput(text if "://" in text else "https://" + text).toString()


def open_href(href, win=None):
    url = QUrl(href)
    if url.isLocalFile():
        p = Path(url.toLocalFile())
        if not p.exists():
            QMessageBox.information(win, APP_NAME, f"Файл не найден:\n{p}")
            return
        if p.suffix.lower() in DOC_EXTS and win is not None and hasattr(win, "open_from_browser"):
            win.open_from_browser(str(p))
            return
    QDesktopServices.openUrl(url)


class TableDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Таблица")
        self.setObjectName("Dialog")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(14)
        title = QLabel("Вставить таблицу")
        title.setObjectName("DialogTitle")
        lay.addWidget(title)
        row = QHBoxLayout()
        self.rows = QSpinBox()
        self.rows.setRange(1, 100)
        self.rows.setValue(3)
        self.cols = QSpinBox()
        self.cols.setRange(1, 20)
        self.cols.setValue(3)
        for lab, sb in (("Строки", self.rows), ("Столбцы", self.cols)):
            col = QVBoxLayout()
            l = QLabel(lab)
            l.setObjectName("Muted")
            col.addWidget(l)
            col.addWidget(sb)
            row.addLayout(col)
        lay.addLayout(row)
        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("Отмена")
        cancel.setObjectName("Ghost")
        ok = QPushButton("Вставить")
        ok.setObjectName("Pill")
        cancel.clicked.connect(self.reject)
        ok.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        lay.addLayout(btns)


# ───────────────────────────── Диалоги ─────────────────────────────
class BaseDialog(QDialog):
    def __init__(self, parent, title, ok_text):
        super().__init__(parent)
        self.setWindowTitle(APP_NAME)
        self.setObjectName("Dialog")
        self.setMinimumWidth(420)
        self.lay = QVBoxLayout(self)
        self.lay.setContentsMargins(24, 22, 24, 18)
        self.lay.setSpacing(10)
        t = QLabel(title)
        t.setObjectName("DialogTitle")
        self.lay.addWidget(t)
        self.lay.addSpacing(4)
        self._ok_text = ok_text

    def field(self, label, text="", placeholder=""):
        l = QLabel(label)
        l.setObjectName("Muted")
        e = QLineEdit(text)
        e.setPlaceholderText(placeholder)
        e.setMinimumHeight(34)
        self.lay.addWidget(l)
        self.lay.addWidget(e)
        return e

    def finish(self):
        self.lay.addSpacing(8)
        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("Отмена")
        cancel.setObjectName("Ghost")
        ok = QPushButton(self._ok_text)
        ok.setObjectName("Pill")
        ok.setDefault(True)
        cancel.clicked.connect(self.reject)
        ok.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        self.lay.addLayout(btns)


class NameDialog(BaseDialog):
    """Название + (по желанию) цвет."""
    def __init__(self, parent, title, ok_text, name="", with_color=False, color=None):
        super().__init__(parent, title, ok_text)
        self.name = self.field("Название", name)
        self.name.selectAll()
        self.color = color or PALETTE[0]
        self.swatches = []
        th = getattr(parent.window(), "th", THEMES["light"]) if parent else THEMES["light"]
        ring = th["text"]
        if with_color:
            l = QLabel("Цвет")
            l.setObjectName("Muted")
            self.lay.addWidget(l)
            row = QHBoxLayout()
            row.setSpacing(8)
            for c in PALETTE:
                b = QToolButton()
                b.setFixedSize(28, 28)
                b.setCheckable(True)
                b.setCursor(Qt.CursorShape.PointingHandCursor)
                b.setStyleSheet(f"QToolButton{{background:{c};border-radius:14px;border:3px solid transparent;}}"
                                f"QToolButton:checked{{border:3px solid {ring};}}"
                                f"QToolButton:hover{{border:3px solid {th['muted']};}}")
                b.setChecked(c.lower() == self.color.lower())
                b.clicked.connect(lambda _=False, col=c: self._pick(col))
                self.swatches.append((c, b))
                row.addWidget(b)
            row.addStretch()
            self.lay.addLayout(row)
        self.finish()

    def _pick(self, c):
        self.color = c
        for col, b in self.swatches:
            b.setChecked(col == c)

    def value(self):
        return safe_name(self.name.text())


class LinkDialog(BaseDialog):
    def __init__(self, parent, title, ok_text, name_label, name="", url="", allow_remove=False):
        super().__init__(parent, title, ok_text)
        self.name = self.field(name_label, name, "Например: Канва — обложки")
        self.url = self.field("Адрес сайта или файла", url, "https://…")
        pick = QPushButton("Выбрать файл на компьютере…")
        pick.setObjectName("Ghost")
        pick.clicked.connect(self._pick_file)
        self.lay.addWidget(pick, 0, Qt.AlignmentFlag.AlignLeft)
        self.removed = False
        if allow_remove:
            rm = QPushButton("Убрать ссылку")
            rm.setObjectName("Ghost")
            rm.clicked.connect(self._remove)
            self.lay.addWidget(rm, 0, Qt.AlignmentFlag.AlignLeft)
        self.finish()
        (self.url if name and not url else self.name if not name else self.url).setFocus()

    def _pick_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "Файл", str(Path.home()))
        if p:
            self.url.setText(QUrl.fromLocalFile(p).toString())
            if not self.name.text().strip():
                self.name.setText(Path(p).stem)

    def _remove(self):
        self.removed = True
        self.accept()

    def values(self):
        url = normalize_url(self.url.text())
        name = self.name.text().strip()
        if not name and url:
            q = QUrl(url)
            name = Path(q.toLocalFile()).stem if q.isLocalFile() else (q.host().removeprefix("www.") or url)
        return name, url


def ask_copy_or_link(parent, n):
    box = QMessageBox(parent)
    box.setWindowTitle(APP_NAME)
    box.setIcon(QMessageBox.Icon.NoIcon)
    what = "файл" if n == 1 else f"файлы ({n})"
    box.setText(f"Как добавить {what}?")
    box.setInformativeText("Копия будет лежать прямо в папке проекта.\n"
                           "Ссылка оставит файл на своём месте — проект будет просто на него указывать.")
    copy = box.addButton("Скопировать сюда", QMessageBox.ButtonRole.AcceptRole)
    link = box.addButton("Прикрепить ссылкой", QMessageBox.ButtonRole.ActionRole)
    box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)
    box.exec()
    c = box.clickedButton()
    return "copy" if c is copy else "link" if c is link else None


# ───────────────────────────── Рабочее пространство ─────────────────────────────
def fonts_dir():
    base = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation) or str(Path.home() / ".folio")
    p = Path(base) / "fonts"
    p.mkdir(parents=True, exist_ok=True)
    return p


def load_user_fonts():
    for f in fonts_dir().iterdir():
        if f.suffix.lower() in (".ttf", ".otf", ".ttc"):
            QFontDatabase.addApplicationFont(str(f))


def workspace_root():
    docs = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation) or str(Path.home())
    p = Path(docs) / APP_NAME
    p.mkdir(parents=True, exist_ok=True)
    return p


def safe_name(name):
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', " ", name).strip().rstrip(".").strip()
    return name[:120] or "Без названия"


def unique_path(folder, stem, ext=""):
    p = Path(folder) / f"{stem}{ext}"
    n = 2
    while p.exists():
        p = Path(folder) / f"{stem} ({n}){ext}"
        n += 1
    return p


def read_url_file(p):
    try:
        for line in Path(p).read_bytes().decode("utf-8-sig", "replace").splitlines():
            if line.lower().startswith("url="):
                return line[4:].strip()
    except OSError:
        pass
    return ""


def write_url_file(p, url):
    enc = bytes(QUrl(url).toEncoded()).decode("ascii", "replace") or url
    Path(p).write_bytes(f"[InternetShortcut]\r\nURL={enc}\r\n".encode("utf-8"))


def _write_hidden(path, text):
    path = Path(path)
    if path.exists():
        with open(path, "r+", encoding="utf-8") as fh:  # скрытый файл в Windows нельзя открыть на «w»
            fh.seek(0)
            fh.write(text)
            fh.truncate()
    else:
        path.write_text(text, encoding="utf-8")
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x02)
            except Exception:
                pass


def folder_color(p):
    try:
        c = json.loads((Path(p) / ".folio").read_text(encoding="utf-8")).get("color")
        if c:
            return c
    except Exception:
        pass
    return PALETTE[zlib.crc32(Path(p).name.encode()) % len(PALETTE)]


def set_folder_color(p, color):
    meta = Path(p) / ".folio"
    data = {}
    try:
        data = json.loads(meta.read_text(encoding="utf-8"))
    except Exception:
        pass
    data["color"] = color
    _write_hidden(meta, json.dumps(data))


def visible(name):
    low = name.lower()
    return not (name.startswith(".") or name.startswith("~$") or low in ("desktop.ini", "thumbs.db"))


def list_dir(d):
    items = []
    try:
        entries = list(os.scandir(d))
    except OSError:
        return items
    for e in entries:
        if not visible(e.name):
            continue
        p = Path(e.path)
        try:
            mtime = e.stat().st_mtime
        except OSError:
            mtime = 0
        if e.is_dir():
            try:
                count = sum(1 for x in os.scandir(p) if visible(x.name))
            except OSError:
                count = 0
            items.append(dict(kind="folder", path=str(p), name=e.name, color=folder_color(p),
                              count=count, mtime=mtime))
        elif p.suffix.lower() == ".url":
            items.append(dict(kind="link", path=str(p), name=p.stem, url=read_url_file(p), mtime=mtime))
        elif p.suffix.lower() in DOC_EXTS:
            items.append(dict(kind="doc", path=str(p), name=p.stem, ext=p.suffix.lower(), mtime=mtime))
        else:
            items.append(dict(kind="file", path=str(p), name=p.stem, ext=p.suffix.lower(), mtime=mtime))
    order = {"folder": 0, "doc": 1, "link": 2, "file": 3}
    items.sort(key=lambda it: (order[it["kind"]], it["name"].casefold()))
    return items


def plural(n, one, few, many):
    n10, n100 = n % 10, n % 100
    if n10 == 1 and n100 != 11:
        return one
    if 2 <= n10 <= 4 and not 12 <= n100 <= 14:
        return few
    return many


def short_date(ts):
    if not ts:
        return ""
    d = datetime.fromtimestamp(ts)
    now = datetime.now()
    if d.date() == now.date():
        return f"сегодня {d:%H:%M}"
    s = f"{d.day} {MONTHS[d.month - 1]}"
    return s if d.year == now.year else f"{s} {d.year}"


_PREVIEW_CACHE = {}


def preview_lines(path, mtime):
    key = (path, mtime)
    if key in _PREVIEW_CACHE:
        return _PREVIEW_CACHE[key]
    lines = []
    p = Path(path)
    ext = p.suffix.lower()
    try:
        if ext == ".docx":
            with zipfile.ZipFile(p) as z:
                xml = z.read("word/document.xml").decode("utf-8", "replace")
            for para in xml.split("</w:p>")[:80]:
                t = "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", para))
                if t.strip():
                    lines.append(html.unescape(t))
                if len(lines) >= 14:
                    break
        else:
            with open(p, "rb") as fh:
                text, _, _ = decode_text(fh.read(4000))
            if ext in HTML_EXTS:
                text = re.sub(r"(?s)<(script|style).*?</\1>", " ", text)
                text = html.unescape(re.sub(r"<[^>]+>", "\n", text))
            for line in text.splitlines():
                line = line.strip()
                if ext in MD_EXTS:
                    line = line.lstrip("#>*- ").strip()
                if line:
                    lines.append(line)
                if len(lines) >= 14:
                    break
    except Exception:
        pass
    _PREVIEW_CACHE[key] = lines
    return lines


# ───────────────────────────── Сетка карточек ─────────────────────────────
class CardDelegate(QStyledItemDelegate):
    W, H = 212, 196

    def __init__(self, win):
        super().__init__()
        self.win = win

    def sizeHint(self, opt, idx):
        return QSize(self.W, self.H)

    def paint(self, p, opt, idx):
        th = self.win.th
        it = idx.data(Qt.ItemDataRole.UserRole)
        if not it:
            return
        p.save()
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)
        r = QRectF(opt.rect).adjusted(7, 7, -7, -7)
        selected = bool(opt.state & QStyle.StateFlag.State_Selected)
        hover = bool(opt.state & QStyle.StateFlag.State_MouseOver)
        drop = idx.row() == getattr(self.win.browser.grid, "drop_row", -1)

        card = QPainterPath()
        rf = getattr(self.win, "rf", 1.0)
        card.addRoundedRect(r, 16 * rf, 16 * rf)
        p.fillPath(card, QColor(th["surface"]))
        if selected or drop:
            p.setPen(QPen(QColor(th["accent"]), 2))
        elif hover:
            p.setPen(QPen(QColor(th["muted"]), 1))
        else:
            p.setPen(QPen(QColor(th["border"]), 1))
        p.drawPath(card)

        prev = QRectF(r.x() + 7, r.y() + 7, r.width() - 14, 116)
        pp = QPainterPath()
        pp.addRoundedRect(prev, 11 * rf, 11 * rf)
        kind = it["kind"]

        if kind == "folder":
            col = QColor(it["color"])
            p.fillPath(pp, col)
            pm = cached_svg("folder", "#FFFFFF", 46)
            p.drawPixmap(int(prev.center().x() - 23), int(prev.center().y() - 30), pm)
            f = QFont(self.win.ui_font)
            f.setPointSizeF(8.5)
            f.setWeight(QFont.Weight.DemiBold)
            p.setFont(f)
            p.setPen(QColor(255, 255, 255, 230))
            n = it.get("count", 0)
            p.drawText(prev.adjusted(0, 0, 0, -12), Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom,
                       f"{n} {plural(n, 'элемент', 'элемента', 'элементов')}" if n else "пусто")
        elif kind == "doc":
            p.fillPath(pp, QColor(th["bg"]))
            p.setClipPath(pp)
            paper = QRectF(prev.x() + 22, prev.y() + 12, prev.width() - 44, prev.height())
            pap = QPainterPath()
            pap.addRoundedRect(paper, 6, 6)
            p.fillPath(pap, QColor(th["page"]))
            p.setPen(QPen(QColor(th["border"]), 1))
            p.drawPath(pap)
            lines = preview_lines(it["path"], it["mtime"])
            y = paper.y() + 10
            for i, line in enumerate(lines):
                f = QFont(self.win.ui_font)
                f.setPixelSize(9 if i == 0 else 7)
                f.setBold(i == 0)
                p.setFont(f)
                p.setPen(QColor(th["text"] if i == 0 else th["muted"]))
                fm = QFontMetrics(f)
                text = fm.elidedText(line, Qt.TextElideMode.ElideRight, int(paper.width() - 18))
                p.drawText(QRectF(paper.x() + 9, y, paper.width() - 18, fm.height()), 0, text)
                y += fm.height() + (4 if i == 0 else 2)
                if y > prev.bottom():
                    break
            if not lines:
                f = QFont(self.win.ui_font)
                f.setPixelSize(8)
                p.setFont(f)
                p.setPen(QColor(th["muted"]))
                p.drawText(paper.adjusted(0, 0, 0, -paper.height() / 2), Qt.AlignmentFlag.AlignCenter, "пустой")
            p.setClipping(False)
            # ближайший дедлайн из заметок
            dl = nearest_deadline(it["path"])
            if dl is not None:
                col = deadline_color(dl) or th["accent"]
                f = QFont(self.win.ui_font)
                f.setPixelSize(10)
                f.setBold(True)
                p.setFont(f)
                txt = deadline_text(dl)
                cw = QFontMetrics(f).horizontalAdvance(txt) + 30
                cr = QRectF(prev.x() + 6, prev.y() + 6, cw, 20)
                cp = QPainterPath()
                cp.addRoundedRect(cr, 10, 10)
                p.fillPath(cp, QColor(col))
                p.drawPixmap(int(cr.x() + 6), int(cr.y() + 3), cached_svg("clock", "#FFFFFF", 14))
                p.setPen(QColor("#FFFFFF"))
                p.drawText(cr.adjusted(22, 0, -6, 0), Qt.AlignmentFlag.AlignVCenter, txt)
            # бейдж формата
            badge = it["ext"].lstrip(".").upper()
            f = QFont(self.win.ui_font)
            f.setPixelSize(9)
            f.setBold(True)
            p.setFont(f)
            bw = QFontMetrics(f).horizontalAdvance(badge) + 12
            br = QRectF(prev.right() - bw - 6, prev.bottom() - 22, bw, 16)
            bp = QPainterPath()
            bp.addRoundedRect(br, 8, 8)
            p.fillPath(bp, QColor(th["accent"]))
            p.setPen(QColor(th["accent_text"]))
            p.drawText(br, Qt.AlignmentFlag.AlignCenter, badge)
        elif kind == "link":
            url = QUrl(it.get("url", ""))
            p.fillPath(pp, QColor(th["soft"]))
            icon = "file" if url.isLocalFile() else "globe"
            p.drawPixmap(int(prev.center().x() - 20), int(prev.center().y() - 30), cached_svg(icon, th["accent"], 40))
            f = QFont(self.win.ui_font)
            f.setPixelSize(11)
            f.setWeight(QFont.Weight.DemiBold)
            p.setFont(f)
            p.setPen(QColor(th["accent"]))
            label = (Path(url.toLocalFile()).name if url.isLocalFile() else url.host().removeprefix("www.")) or "ссылка"
            fm = QFontMetrics(f)
            p.drawText(prev.adjusted(10, 0, -10, -14), Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom,
                       fm.elidedText(label, Qt.TextElideMode.ElideMiddle, int(prev.width() - 20)))
            p.drawPixmap(int(prev.right() - 26), int(prev.y() + 8), cached_svg("link", th["accent"], 18))
        else:
            p.fillPath(pp, QColor(th["bg"]))
            p.drawPixmap(int(prev.center().x() - 20), int(prev.center().y() - 28), cached_svg("file", th["muted"], 40))
            f = QFont(self.win.ui_font)
            f.setPixelSize(11)
            f.setBold(True)
            p.setFont(f)
            p.setPen(QColor(th["muted"]))
            p.drawText(prev.adjusted(0, 0, 0, -14), Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom,
                       it.get("ext", "").lstrip(".").upper() or "ФАЙЛ")

        # подпись
        f = QFont(self.win.ui_font)
        f.setPointSizeF(10)
        f.setWeight(QFont.Weight.DemiBold)
        p.setFont(f)
        p.setPen(QColor(th["text"]))
        fm = QFontMetrics(f)
        name_r = QRectF(r.x() + 12, prev.bottom() + 10, r.width() - 24, fm.height())
        p.drawText(name_r, 0, fm.elidedText(it["name"], Qt.TextElideMode.ElideRight, int(name_r.width())))
        f2 = QFont(self.win.ui_font)
        f2.setPointSizeF(8.5)
        p.setFont(f2)
        p.setPen(QColor(th["muted"]))
        if kind == "folder":
            sub = "Проект" if Path(it["path"]).parent == self.win.browser.root else "Папка"
        elif kind == "link":
            sub = "Ссылка на файл" if QUrl(it.get("url", "")).isLocalFile() else "Ссылка"
        else:
            sub = short_date(it["mtime"])
        p.drawText(QRectF(name_r.x(), name_r.bottom() + 2, name_r.width(), 18), 0, sub)
        p.restore()


class Grid(QListWidget):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.drop_row = -1
        self.setViewMode(QListView.ViewMode.IconMode)
        self.setResizeMode(QListView.ResizeMode.Adjust)
        self.setMovement(QListView.Movement.Static)
        self.setUniformItemSizes(True)
        self.setGridSize(QSize(CardDelegate.W, CardDelegate.H))
        self.setSpacing(0)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(False)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.CopyAction)
        self.setMouseTracking(True)
        self.viewport().setAttribute(Qt.WidgetAttribute.WA_Hover)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.verticalScrollBar().setSingleStep(24)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.setObjectName("Grid")

    def mimeTypes(self):
        return [INTERNAL_MIME, "text/uri-list"]

    def mimeData(self, items):
        md = QMimeData()
        paths = [i.data(Qt.ItemDataRole.UserRole)["path"] for i in items]
        md.setUrls([QUrl.fromLocalFile(x) for x in paths])
        md.setData(INTERNAL_MIME, QByteArray("\n".join(paths).encode("utf-8")))
        return md

    def supportedDropActions(self):
        return Qt.DropAction.CopyAction | Qt.DropAction.MoveAction | Qt.DropAction.LinkAction

    def _target_row(self, pos):
        item = self.itemAt(pos)
        if item and item.data(Qt.ItemDataRole.UserRole)["kind"] == "folder" and not item.isSelected():
            return self.row(item)
        return -1

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(INTERNAL_MIME) or e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        row = self._target_row(e.position().toPoint())
        if row != self.drop_row:
            self.drop_row = row
            self.viewport().update()
        internal = e.mimeData().hasFormat(INTERNAL_MIME)
        if internal and row < 0:
            e.ignore()  # перетаскивание внутри той же папки ничего не делает
            return
        e.setDropAction(Qt.DropAction.CopyAction)
        e.accept()

    def dragLeaveEvent(self, e):
        self.drop_row = -1
        self.viewport().update()
        super().dragLeaveEvent(e)

    def dropEvent(self, e):
        row = self._target_row(e.position().toPoint())
        self.drop_row = -1
        target = Path(self.item(row).data(Qt.ItemDataRole.UserRole)["path"]) if row >= 0 else self.page.cwd
        md = e.mimeData()
        e.setDropAction(Qt.DropAction.CopyAction)
        e.accept()
        if md.hasFormat(INTERNAL_MIME):
            paths = bytes(md.data(INTERNAL_MIME)).decode("utf-8").split("\n")
            QTimer.singleShot(0, lambda: self.page.move_into(paths, target))
        elif md.hasUrls():
            urls = md.urls()
            title = md.text() if md.hasText() else ""
            QTimer.singleShot(0, lambda: self.page.add_external(urls, target, title))

    def keyPressEvent(self, e):
        k = e.key()
        if k in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if self.currentItem():
                self.page.open_item(self.currentItem())
            return
        if k == Qt.Key.Key_Delete:
            self.page.delete_selected()
            return
        if k == Qt.Key.Key_F2:
            if self.currentItem():
                self.page.rename_item(self.currentItem())
            return
        if k == Qt.Key.Key_Backspace:
            self.page.go_up()
            return
        super().keyPressEvent(e)


class CrumbButton(QPushButton):
    """Элемент «пути» сверху. На него можно перетащить файл, чтобы переместить его на уровень выше."""
    def __init__(self, text, path, page, current=False):
        super().__init__(text)
        self.path = Path(path)
        self.page = page
        self.setObjectName("CrumbCur" if current else "Crumb")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAcceptDrops(not current)
        self.clicked.connect(lambda: page.navigate(self.path))

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(INTERNAL_MIME):
            e.acceptProposedAction()
            self.setProperty("drop", True)
            self.style().unpolish(self)
            self.style().polish(self)

    def dragLeaveEvent(self, e):
        self.setProperty("drop", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, e):
        self.dragLeaveEvent(e)
        paths = bytes(e.mimeData().data(INTERNAL_MIME)).decode("utf-8").split("\n")
        e.setDropAction(Qt.DropAction.CopyAction)
        e.accept()
        QTimer.singleShot(0, lambda: self.page.move_into(paths, self.path))


# ───────────────────────────── Экран проектов ─────────────────────────────
class BrowserPage(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.root = workspace_root()
        self.cwd = self.root
        self.setObjectName("Browser")

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        bar = QWidget()
        bar.setObjectName("AppBar")
        bar.setFixedHeight(64)
        h = QHBoxLayout(bar)
        h.setContentsMargins(16, 10, 16, 6)
        h.setSpacing(4)
        logo = QLabel()
        logo.setFixedSize(36, 36)
        logo.setPixmap(app_icon().pixmap(32, 32))
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        h.addWidget(logo)
        h.addSpacing(8)
        self.crumbs = QHBoxLayout()
        self.crumbs.setSpacing(0)
        h.addLayout(self.crumbs)
        h.addStretch()
        h.addWidget(win._btn("Открыть файл с компьютера  (Ctrl+O)", ("svg", "open"), win.open_dialog))
        self.theme_btn = win._btn("Тёмная / светлая тема", ("svg", "moon"), win.toggle_theme)
        h.addWidget(self.theme_btn)
        h.addWidget(win._btn("Внешний вид: цвета, страница, шрифт", ("svg", "tune"), win.open_settings))
        h.addSpacing(8)
        self.update_btn = QPushButton("Обновить")
        self.update_btn.setObjectName("Ghost")
        self.update_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.update_btn.hide()
        self.update_btn.clicked.connect(lambda: win.updater.ask_install())
        h.addWidget(self.update_btn)
        h.addSpacing(6)
        self.create_btn = QPushButton("  Создать")
        self.create_btn.setObjectName("Pill")
        self.create_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.create_btn.clicked.connect(self.show_create_menu)
        h.addWidget(self.create_btn)
        v.addWidget(bar)

        body = QWidget()
        body.setObjectName("EditorArea")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(24, 8, 24, 0)
        hrow = QHBoxLayout()
        self.heading = QLabel()
        self.heading.setObjectName("PageHeading")
        self.toast = QLabel("")
        self.toast.setObjectName("Toast")
        self.toast_timer = QTimer(self)
        self.toast_timer.setSingleShot(True)
        self.toast_timer.timeout.connect(lambda: self.toast.setText(""))
        hrow.addWidget(self.heading)
        hrow.addStretch()
        hrow.addWidget(self.toast)
        bl.addLayout(hrow)
        bl.addSpacing(6)
        self.grid = Grid(self)
        self.grid.setItemDelegate(CardDelegate(win))
        self.grid.itemDoubleClicked.connect(self.open_item)
        self.grid.customContextMenuRequested.connect(self.context_menu)
        bl.addWidget(self.grid, 1)
        self.empty = QLabel()
        self.empty.setObjectName("Empty")
        self.empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty.setWordWrap(True)
        bl.addWidget(self.empty, 1)
        v.addWidget(body, 1)

        self.watcher = QFileSystemWatcher(self)
        self.watcher.directoryChanged.connect(lambda _: self.refresh_timer.start())
        self.refresh_timer = QTimer(self)
        self.refresh_timer.setSingleShot(True)
        self.refresh_timer.setInterval(300)
        self.refresh_timer.timeout.connect(self.refresh)

    # ── навигация ──
    def at_root(self):
        return self.cwd == self.root

    def navigate(self, path):
        path = Path(path)
        if not path.exists() or not path.is_dir():
            path = self.root
        try:
            path.relative_to(self.root)
        except ValueError:
            path = self.root
        self.cwd = path
        if self.watcher.directories():
            self.watcher.removePaths(self.watcher.directories())
        self.watcher.addPath(str(path))
        self._build_crumbs()
        self.refresh()

    def go_up(self):
        if not self.at_root():
            self.navigate(self.cwd.parent)

    def _build_crumbs(self):
        while self.crumbs.count():
            w = self.crumbs.takeAt(0).widget()
            if w:
                w.hide()
                w.setParent(None)
                w.deleteLater()
        chain = [self.root]
        if not self.at_root():
            rel = self.cwd.relative_to(self.root)
            acc = self.root
            for part in rel.parts:
                acc = acc / part
                chain.append(acc)
        # длинные пути сокращаем: Проекты › … › предпоследняя › текущая
        shown = chain if len(chain) <= 4 else [chain[0], None] + chain[-2:]
        for i, c in enumerate(shown):
            if i:
                s = QLabel("›")
                s.setObjectName("CrumbSep")
                self.crumbs.addWidget(s)
            if c is None:
                b = CrumbButton("…", chain[-3], self)
            else:
                b = CrumbButton("Проекты" if c == self.root else c.name, c, self, current=(c == self.cwd))
            self.crumbs.addWidget(b)
        self.heading.setText("Мои проекты" if self.at_root() else self.cwd.name)
        self.create_btn.setText("  Новый проект" if self.at_root() else "  Создать")

    def refresh(self):
        selected = {i.data(Qt.ItemDataRole.UserRole)["path"] for i in self.grid.selectedItems()}
        if not self.cwd.exists():
            self.navigate(self.root)
            return
        self.grid.clear()
        items = list_dir(self.cwd)
        for it in items:
            li = QListWidgetItem()
            li.setData(Qt.ItemDataRole.UserRole, it)
            li.setSizeHint(QSize(CardDelegate.W, CardDelegate.H))
            li.setToolTip(it.get("url") or it["path"])
            self.grid.addItem(li)
            if it["path"] in selected:
                li.setSelected(True)
        empty = not items
        self.grid.setVisible(not empty)
        self.empty.setVisible(empty)
        if self.at_root():
            self.empty.setText("Здесь будут твои проекты\n\nНажми «Новый проект» — это как папка в Канве:\n"
                               "внутри будут документы, папки и ссылки.")
        else:
            self.empty.setText("Пока пусто\n\nНажми «Создать», чтобы добавить документ, папку или ссылку,\n"
                               "или перетащи сюда файлы из проводника.")

    def select_path(self, path):
        self.grid.clearSelection()
        for i in range(self.grid.count()):
            it = self.grid.item(i)
            if it.data(Qt.ItemDataRole.UserRole)["path"] == str(path):
                it.setSelected(True)
                self.grid.setCurrentItem(it)
                self.grid.scrollToItem(it)
                break

    # ── создание ──
    def create_menu(self):
        m = styled_menu(self)
        if self.at_root():
            m.addAction(make_icon(("svg", "folder"), self.win.th), "Новый проект", self.new_folder)
        else:
            m.addAction(make_icon(("svg", "new"), self.win.th), "Документ", self.new_doc)
            m.addAction(make_icon(("svg", "folder"), self.win.th), "Папка", self.new_folder)
            m.addAction(make_icon(("svg", "link"), self.win.th), "Ссылка", self.new_link)
            m.addSeparator()
            m.addAction(make_icon(("svg", "open"), self.win.th), "Добавить файлы с компьютера…", self.add_files)
        return m

    def show_create_menu(self):
        if self.at_root():
            self.new_folder()
            return
        m = self.create_menu()
        m.exec(self.create_btn.mapToGlobal(self.create_btn.rect().bottomLeft()))

    def new_folder(self):
        project = self.at_root()
        dlg = NameDialog(self, "Новый проект" if project else "Новая папка", "Создать",
                         "Новый проект" if project else "Новая папка", with_color=True,
                         color=PALETTE[len(list_dir(self.cwd)) % len(PALETTE)])
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        p = unique_path(self.cwd, dlg.value())
        try:
            p.mkdir(parents=True)
            set_folder_color(p, dlg.color)
        except OSError as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось создать папку:\n{e}")
            return
        if project:
            self.navigate(p)
        else:
            self.refresh()
            self.select_path(p)

    def new_doc(self):
        dlg = NameDialog(self, "Новый документ", "Создать", "Без названия")
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        p = unique_path(self.cwd, dlg.value(), ".docx")
        try:
            save_docx(QTextDocument(), str(p))
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось создать документ:\n{e}")
            return
        self.refresh()
        self.win.open_from_browser(str(p))

    def new_link(self):
        clip = QGuiApplication.clipboard().text().strip()
        url = clip if re.match(r"^https?://\S+$", clip) else ""
        dlg = LinkDialog(self, "Новая ссылка", "Добавить", "Название (так будет видно на карточке)", url=url)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        name, url = dlg.values()
        if not url:
            return
        p = unique_path(self.cwd, safe_name(name), ".url")
        write_url_file(p, url)
        self.refresh()
        self.select_path(p)

    def add_files(self):
        start = self.win.settings.value("last_dir", str(Path.home() / "Documents"))
        paths, _ = QFileDialog.getOpenFileNames(self, "Добавить файлы", start)
        if paths:
            self.win.settings.setValue("last_dir", str(Path(paths[0]).parent))
            self.add_external([QUrl.fromLocalFile(x) for x in paths], self.cwd)

    def add_external(self, urls, target, title=""):
        target = Path(target)
        local = [Path(u.toLocalFile()) for u in urls if u.isLocalFile()]
        web = [u for u in urls if not u.isLocalFile() and u.scheme() in ("http", "https")]
        last = None
        for u in web:  # ссылку перетащили из браузера — делаем карточку
            name = title.strip() if title and not title.strip().startswith("http") else u.host().removeprefix("www.")
            last = unique_path(target, safe_name(name or "Ссылка"), ".url")
            write_url_file(last, u.toString())
        local = [p for p in local if p.exists()]
        if local:
            if target == self.root:
                QMessageBox.information(self, APP_NAME, "Сначала открой проект — файлы добавляются внутрь проекта.")
                return
            mode = ask_copy_or_link(self, len(local))
            if not mode:
                return
            for src in local:
                try:
                    if mode == "link":
                        last = unique_path(target, safe_name(src.stem if src.is_file() else src.name), ".url")
                        write_url_file(last, QUrl.fromLocalFile(str(src)).toString())
                    elif src.is_dir():
                        last = unique_path(target, src.name)
                        shutil.copytree(src, last)
                    else:
                        last = unique_path(target, src.stem, src.suffix)
                        shutil.copy2(src, last)
                except Exception as e:
                    QMessageBox.warning(self, APP_NAME, f"Не получилось добавить «{src.name}»:\n{e}")
        self.refresh()
        if last and last.parent == self.cwd:
            self.select_path(last)

    # ── действия с карточками ──
    def item_data(self, item):
        return item.data(Qt.ItemDataRole.UserRole)

    def open_item(self, item):
        it = self.item_data(item)
        if it["kind"] == "folder":
            self.navigate(it["path"])
        elif it["kind"] == "doc":
            self.win.open_from_browser(it["path"])
        elif it["kind"] == "link":
            if it.get("url"):
                open_href(it["url"], self.win)
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(it["path"]))

    def rename_item(self, item):
        it = self.item_data(item)
        p = Path(it["path"])
        is_dir = it["kind"] == "folder"
        dlg = NameDialog(self, "Переименовать", "Сохранить", p.name if is_dir else p.stem)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        new = p.with_name(dlg.value() + ("" if is_dir else p.suffix))
        if new == p:
            return
        if new.exists() and new.name.lower() != p.name.lower():
            QMessageBox.information(self, APP_NAME, "Здесь уже есть элемент с таким названием.")
            return
        try:
            p.rename(new)
            move_with_sidecar(p, new)
        except OSError as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось переименовать:\n{e}")
            return
        self.win.path_moved(p, new)
        self.refresh()
        self.select_path(new)

    def edit_link(self, item):
        it = self.item_data(item)
        p = Path(it["path"])
        dlg = LinkDialog(self, "Изменить ссылку", "Сохранить", "Название (так будет видно на карточке)",
                         name=p.stem, url=it.get("url", ""))
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        name, url = dlg.values()
        if not url:
            return
        new = p.with_name(safe_name(name) + ".url")
        if new != p and new.exists():
            new = unique_path(p.parent, safe_name(name), ".url")
        write_url_file(p, url)
        if new != p:
            p.rename(new)
        self.refresh()
        self.select_path(new)

    def recolor(self, item, color):
        set_folder_color(self.item_data(item)["path"], color.name() if isinstance(color, QColor) else color)
        self.refresh()

    def reveal(self, path):
        path = Path(path)
        if sys.platform == "win32":
            subprocess.Popen(f'explorer /select,"{path}"')
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.parent)))

    def delete_selected(self):
        items = self.grid.selectedItems()
        if not items:
            return
        names = [self.item_data(i)["name"] for i in items]
        what = f"«{names[0]}»" if len(names) == 1 else f"{len(names)} {plural(len(names), 'элемент', 'элемента', 'элементов')}"
        box = QMessageBox(self)
        box.setWindowTitle(APP_NAME)
        box.setIcon(QMessageBox.Icon.NoIcon)
        box.setText(f"Удалить {what}?")
        box.setInformativeText("Удалённое попадёт в Корзину Windows — его можно будет восстановить.")
        yes = box.addButton("Удалить", QMessageBox.ButtonRole.DestructiveRole)
        box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is not yes:
            return
        for i in items:
            path = self.item_data(i)["path"]
            sc = sidecar_path(path)
            if Path(path).is_file() and sc.exists():
                QFile.moveToTrash(str(sc))
            res_ = QFile.moveToTrash(path)
            ok = res_[0] if isinstance(res_, tuple) else bool(res_)
            if not ok:
                QMessageBox.warning(self, APP_NAME, f"Не получилось удалить «{Path(path).name}».\n"
                                                    "Возможно, файл открыт в другой программе.")
        self.refresh()

    def move_into(self, paths, target):
        target = Path(target)
        moved = None
        for s in paths:
            src = Path(s)
            if not src.exists() or src.parent == target or src == target:
                continue
            try:
                target.relative_to(src)  # нельзя положить папку саму в себя
                continue
            except ValueError:
                pass
            dst = target / src.name
            if dst.exists():
                dst = unique_path(target, src.stem if src.is_file() else src.name, src.suffix if src.is_file() else "")
            try:
                shutil.move(str(src), str(dst))
                move_with_sidecar(src, dst)
                self.win.path_moved(src, dst)
                moved = dst
            except Exception as e:
                QMessageBox.warning(self, APP_NAME, f"Не получилось переместить «{src.name}»:\n{e}")
        self.refresh()
        if moved:
            self.win.show_browser_toast(f"Перемещено в «{'Проекты' if target == self.root else target.name}»")

    def context_menu(self, pos):
        item = self.grid.itemAt(pos)
        th = self.win.th
        if item is None:
            m = self.create_menu()
            m.addSeparator()
            m.addAction("Показать в проводнике", lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.cwd))))
            m.exec(self.grid.viewport().mapToGlobal(pos))
            return
        if not item.isSelected():
            self.grid.clearSelection()
            item.setSelected(True)
            self.grid.setCurrentItem(item)
        it = self.item_data(item)
        m = styled_menu(self)
        m.addAction("Открыть", lambda: self.open_item(item))
        if it["kind"] == "link":
            m.addAction(make_icon(("svg", "link"), th), "Изменить ссылку…", lambda: self.edit_link(item))
            m.addAction("Копировать адрес", lambda: QGuiApplication.clipboard().setText(it.get("url", "")))
        else:
            m.addAction("Переименовать", lambda: self.rename_item(item)).setShortcut("F2")
        if it["kind"] == "folder":
            sm = SwatchMenu(m, colors=PALETTE, reset=False)
            sm.setTitle("Цвет")
            sm.picked.connect(lambda c: self.recolor(item, c))
            m.addMenu(sm)
        if it["kind"] == "doc":
            m.addMenu(export_menu(m, lambda f, p=it["path"]: self.win.export_file(p, f), th))
        m.addAction("Показать в проводнике", lambda: self.reveal(it["path"]))
        m.addSeparator()
        m.addAction("Удалить", self.delete_selected).setShortcut("Del")
        m.exec(self.grid.viewport().mapToGlobal(pos))


# ───────────────────────────── Настройки внешнего вида ─────────────────────────────
PAGE_PRESETS = {
    "light": ["", "#FBF7EE", "#F6F6F2", "#EEF3FF", "#EEF7F0", "#FFF1F3"],
    "dark": ["", "#1E1E1E", "#1B1D2A", "#1F2421", "#26201C", "#000000"],
}
TEXT_PRESETS = {
    "light": ["", "#000000", "#3C4043", "#4A3A2A", "#1F3A5F"],
    "dark": ["", "#FFFFFF", "#C9CDD4", "#F1E6CF", "#CFE3FF"],
}


class SettingsDialog(QDialog):
    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.s = win.settings
        self.setWindowTitle(APP_NAME)
        self.setObjectName("Dialog")
        self.setMinimumWidth(560)
        from PySide6.QtWidgets import QScrollArea
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.Shape.NoFrame)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        inner = QWidget()
        inner.setObjectName("Dialog")
        sc.setWidget(inner)
        outer.addWidget(sc)
        avail = QGuiApplication.primaryScreen().availableGeometry().height() if QGuiApplication.primaryScreen() else 900
        self.resize(600, min(840, avail - 80))
        lay = QVBoxLayout(inner)
        lay.setContentsMargins(26, 22, 26, 18)
        lay.setSpacing(8)
        t = QLabel("Внешний вид")
        t.setObjectName("DialogTitle")
        lay.addWidget(t)

        lay.addWidget(self._section("ТЕМА"))
        row = QHBoxLayout()
        row.setSpacing(8)
        self.mode_btns = {}
        for key, label in (("light", "Светлая"), ("dark", "Тёмная"), ("system", "Как в Windows")):
            b = QPushButton(label)
            b.setObjectName("Seg")
            b.setCheckable(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.clicked.connect(lambda _=False, k=key: self.set_mode(k))
            self.mode_btns[key] = b
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)

        lay.addWidget(self._section("ЦВЕТ ИНТЕРФЕЙСА"))
        self.accent_row = QHBoxLayout()
        self.accent_row.setSpacing(8)
        lay.addLayout(self.accent_row)
        self.tint = QCheckBox("Слегка подкрасить фон в этот цвет")
        self.tint.setCursor(Qt.CursorShape.PointingHandCursor)
        self.tint.toggled.connect(self.set_tint)
        lay.addWidget(self.tint)
        self.glass = QCheckBox("Стеклянная панель инструментов (размытие, как в iPhone)")
        self.glass.setCursor(Qt.CursorShape.PointingHandCursor)
        self.glass.toggled.connect(self.set_glass)
        lay.addWidget(self.glass)

        self.page_title = self._section("")
        lay.addWidget(self.page_title)
        self.page_row = QHBoxLayout()
        self.page_row.setSpacing(8)
        lay.addLayout(self.page_row)

        self.text_title = self._section("")
        lay.addWidget(self.text_title)
        self.text_row = QHBoxLayout()
        self.text_row.setSpacing(8)
        lay.addLayout(self.text_row)

        lay.addWidget(self._section("ИНТЕРФЕЙС"))
        from PySide6.QtWidgets import QSlider
        def slider_row(label, lo, hi, val, key, scale=1.0, suffix=""):
            row = QHBoxLayout()
            l = QLabel(label)
            l.setFixedWidth(170)
            sl = QSlider(Qt.Orientation.Horizontal)
            sl.setRange(lo, hi)
            sl.setValue(int(round(val)))
            vl = QLabel(f"{int(round(val))}{suffix}")
            vl.setObjectName("Muted")
            vl.setFixedWidth(52)
            def changed(v):
                vl.setText(f"{v}{suffix}")
                self.s.setValue(key, v * scale)
                self.win.apply_theme()
            sl.valueChanged.connect(changed)
            row.addWidget(l)
            row.addWidget(sl, 1)
            row.addWidget(vl)
            lay.addLayout(row)
            return sl
        self.sl_round = slider_row("Скругление углов", 0, 160, float(self.s.value("ui_round", 1.0)) * 100,
                                   "ui_round", 0.01, "%")
        self.sl_font = slider_row("Размер шрифта меню", 8, 13, float(self.s.value("ui_font", 10)), "ui_font", 1, " pt")
        self.sl_page = slider_row("Ширина страницы", 600, 1400, float(self.s.value("page_width", 880)),
                                  "page_width", 1, " px")
        self.sl_tint = slider_row("Стекло: плотность", 5, 90, float(self.s.value("glass_tint", 0.40)) * 100,
                                  "glass_tint", 0.01, "%")
        self.sl_blur = slider_row("Стекло: размытие", 2, 30, float(self.s.value("glass_blur", 12)), "glass_blur", 1)
        self.snap_cb = QCheckBox("Притягивать стикеры к сетке и друг к другу")
        self.snap_cb.setChecked(self.s.value("snap", "true") in (True, "true"))
        self.snap_cb.toggled.connect(lambda on: self.s.setValue("snap", "true" if on else "false"))
        lay.addWidget(self.snap_cb)
        self.insp_cb = QCheckBox("Открывать панель свойств, когда выбираю стикер или стрелку")
        self.insp_cb.setChecked(self.s.value("inspector_auto", "true") in (True, "true"))
        self.insp_cb.toggled.connect(lambda on: self.s.setValue("inspector_auto", "true" if on else "false"))
        lay.addWidget(self.insp_cb)

        lay.addWidget(self._section("ШРИФТ ДОКУМЕНТОВ"))
        frow = QHBoxLayout()
        self.font = QFontComboBox()
        self.font.setCurrentFont(QFont(DEFAULT_FONT))
        self.font.setMinimumWidth(220)
        self.font.currentFontChanged.connect(self.set_font)
        self.size = QComboBox()
        self.size.addItems([str(x) for x in (10, 11, 12, 13, 14, 16, 18)])
        self.size.setCurrentText(f"{DEFAULT_SIZE:g}")
        self.size.currentTextChanged.connect(self.set_size)
        for w in (self.font, self.size):
            w.setStyleSheet(f"background: {win.th['bg']}; border: 1px solid {win.th['border']};")
        frow.addWidget(self.font)
        frow.addWidget(self.size)
        addf = QPushButton("Добавить шрифт из файла…")
        addf.setObjectName("Ghost")
        addf.clicked.connect(self.add_font)
        frow.addWidget(addf)
        frow.addStretch()
        lay.addLayout(frow)
        hint = QLabel("Используется для текста без своего оформления и в новых документах.")
        hint.setObjectName("Muted")
        lay.addWidget(hint)

        lay.addWidget(self._section("ОБНОВЛЕНИЯ"))
        urow = QHBoxLayout()
        ver = QLabel(f"Folio {APP_VERSION}")
        ver.setObjectName("Muted")
        chk = QPushButton("Проверить обновления")
        chk.setObjectName("Ghost")
        chk.clicked.connect(lambda: self.win.updater.check(True))
        urow.addWidget(ver)
        urow.addStretch()
        urow.addWidget(chk)
        lay.addLayout(urow)
        self.auto_upd = QCheckBox("Проверять обновления при запуске")
        self.auto_upd.setChecked(self.s.value("auto_update", "true") in (True, "true"))
        self.auto_upd.toggled.connect(lambda on: self.s.setValue("auto_update", "true" if on else "false"))
        lay.addWidget(self.auto_upd)

        lay.addWidget(self._section("РЕЗЕРВНАЯ КОПИЯ НАСТРОЕК"))
        brow = QHBoxLayout()
        exp = QPushButton("Сохранить в файл…")
        exp.setObjectName("Ghost")
        exp.clicked.connect(self.export_settings)
        imp = QPushButton("Загрузить из файла…")
        imp.setObjectName("Ghost")
        imp.clicked.connect(self.import_settings)
        brow.addWidget(exp)
        brow.addWidget(imp)
        brow.addStretch()
        lay.addLayout(brow)
        note = QLabel("Настройки хранятся отдельно от программы — обновления их не сбрасывают.")
        note.setObjectName("Muted")
        lay.addWidget(note)

        lay.addSpacing(10)
        btns = QHBoxLayout()
        reset = QPushButton("Сбросить всё")
        reset.setObjectName("Ghost")
        reset.clicked.connect(self.reset_all)
        done = QPushButton("Готово")
        done.setObjectName("Pill")
        done.setDefault(True)
        done.clicked.connect(self.accept)
        btns.addWidget(reset)
        btns.addStretch()
        btns.addWidget(done)
        lay.addLayout(btns)
        self.rebuild()

    def _section(self, text):
        l = QLabel(text)
        l.setObjectName("Section")
        return l

    def _swatch(self, color, checked, slot, tip="", default_fill=None):
        th = self.win.th
        b = QToolButton()
        b.setFixedSize(30, 30)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setToolTip(tip)
        fill = color or default_fill
        ring = th["accent"] if checked else th["border"]
        width = 3 if checked else 1
        b.setStyleSheet(f"QToolButton{{background:{fill};border-radius:15px;border:{width}px solid {ring};}}"
                        f"QToolButton:hover{{border:3px solid {th['muted']};}}")
        if not color:  # «по умолчанию» — с маленькой галочкой-точкой
            b.setStyleSheet(b.styleSheet() + f"QToolButton{{color:{th['muted']};font-size:7pt;font-weight:700;}}")
            b.setText("авто")
            b.setFixedSize(42, 30)
        b.clicked.connect(slot)
        return b

    def _plus(self, slot, tip):
        b = QToolButton()
        b.setFixedSize(30, 30)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setToolTip(tip)
        b.setIcon(QIcon(svg_pixmap(SVG["plus"], self.win.th["text"], size=16)))
        b.setStyleSheet(f"QToolButton{{border:1px dashed {self.win.th['muted']};border-radius:15px;}}")
        b.clicked.connect(slot)
        return b

    @staticmethod
    def _clear(layout):
        while layout.count():
            w = layout.takeAt(0).widget()
            if w:
                w.hide()
                w.setParent(None)
                w.deleteLater()

    def rebuild(self):
        s, win = self.s, self.win
        th = win.th
        for k, b in self.mode_btns.items():
            b.setChecked(k == win.theme_mode)
        # акцент
        self._clear(self.accent_row)
        cur = (s.value("accent", "") or THEMES["light"]["accent"]).lower()
        for c in PALETTE + ["#111827"]:
            self.accent_row.addWidget(self._swatch(c, c.lower() == cur, lambda _=False, col=c: self.set_accent(col)))
        custom = cur not in [c.lower() for c in PALETTE + ["#111827"]]
        if custom:
            self.accent_row.addWidget(self._swatch(cur, True, lambda: None, "Свой цвет"))
        self.accent_row.addWidget(self._plus(self.pick_accent, "Выбрать свой цвет"))
        self.accent_row.addStretch()
        self.tint.blockSignals(True)
        self.tint.setChecked(s.value("tint", "false") in (True, "true"))
        self.tint.blockSignals(False)
        self.glass.blockSignals(True)
        self.glass.setChecked(s.value("glass", "true") in (True, "true"))
        self.glass.blockSignals(False)
        # страница и текст — отдельно для светлой и тёмной темы
        mode = win.theme
        which = "светлой" if mode == "light" else "тёмной"
        self.page_title.setText(f"ЦВЕТ СТРАНИЦЫ  ·  для {which} темы")
        self.text_title.setText(f"ЦВЕТ ТЕКСТА  ·  для {which} темы")
        base = THEMES[mode]
        for row, key, presets, default, picker in (
                (self.page_row, f"page_{mode}", PAGE_PRESETS[mode], base["page"], self.pick_page),
                (self.text_row, f"text_{mode}", TEXT_PRESETS[mode], base["text"], self.pick_text)):
            self._clear(row)
            curv = (s.value(key, "") or "").lower()
            for c in presets:
                row.addWidget(self._swatch(c, c.lower() == curv, lambda _=False, k=key, col=c: self.set_value(k, col),
                                           "По умолчанию" if not c else c, default_fill=default))
            if curv and curv not in [c.lower() for c in presets]:
                row.addWidget(self._swatch(curv, True, lambda: None, "Свой цвет"))
            row.addWidget(self._plus(picker, "Выбрать свой цвет"))
            row.addStretch()

    # ── изменения применяются сразу ──
    def _apply(self):
        self.win.apply_theme()
        self.rebuild()

    def set_mode(self, k):
        self.win.theme_mode = k
        self.s.setValue("theme", k)
        self.win.theme = self.win.resolve_theme()
        self._apply()

    def set_accent(self, c):
        self.s.setValue("accent", c)
        self._apply()

    def set_tint(self, on):
        self.s.setValue("tint", "true" if on else "false")
        self._apply()

    SETTINGS_KEYS = ("theme", "accent", "tint", "glass", "glass_tint", "glass_blur", "page_light", "page_dark",
                     "text_light", "text_dark", "doc_font", "doc_size", "ui_round", "ui_font", "page_width",
                     "snap", "inspector_auto", "recent_colors")

    def export_settings(self):
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить настройки", str(Path.home() / "Documents" / "Folio-настройки.json"),
                                              "Настройки Folio (*.json)")
        if not path:
            return
        data = {k: self.s.value(k) for k in self.SETTINGS_KEYS if self.s.contains(k)}
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.win.show_toast("Настройки сохранены ✓")

    def import_settings(self):
        path, _ = QFileDialog.getOpenFileName(self, "Загрузить настройки", str(Path.home() / "Documents"),
                                              "Настройки Folio (*.json)")
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось прочитать файл:\n{e}")
            return
        for k, v in data.items():
            if k in self.SETTINGS_KEYS:
                self.s.setValue(k, v)
        global DEFAULT_FONT, DEFAULT_SIZE
        DEFAULT_FONT = self.s.value("doc_font", DEFAULT_FONT) or DEFAULT_FONT
        DEFAULT_SIZE = float(self.s.value("doc_size", DEFAULT_SIZE))
        self.win.theme_mode = self.s.value("theme", "system")
        self.win.theme = self.win.resolve_theme()
        self._apply()

    def set_glass(self, on):
        self.s.setValue("glass", "true" if on else "false")
        self._apply()

    def add_font(self):
        self.win.add_font_file()
        self.font.setFontFilters(self.font.fontFilters())

    def set_value(self, key, c):
        self.s.setValue(key, c)
        self._apply()

    def _pick(self, initial):
        c = QColorDialog.getColor(QColor(initial), self, "Свой цвет")
        return c.name() if c.isValid() else None

    def pick_accent(self):
        c = self._pick(self.win.th["accent"])
        if c:
            self.set_accent(c)

    def pick_page(self):
        c = self._pick(self.win.th["page"])
        if c:
            self.set_value(f"page_{self.win.theme}", c)

    def pick_text(self):
        c = self._pick(self.win.th["doc_text"])
        if c:
            self.set_value(f"text_{self.win.theme}", c)

    def set_font(self, f):
        global DEFAULT_FONT
        DEFAULT_FONT = f.family()
        self.s.setValue("doc_font", DEFAULT_FONT)
        self.win.apply_theme()

    def set_size(self, t):
        global DEFAULT_SIZE
        try:
            DEFAULT_SIZE = float(t)
        except ValueError:
            return
        self.s.setValue("doc_size", DEFAULT_SIZE)
        self.win.apply_theme()

    def reset_all(self):
        for k in ("ui_round", "ui_font", "page_width", "glass_tint", "glass_blur", "snap", "inspector_auto",
                  "accent", "tint", "glass", "page_light", "page_dark", "text_light", "text_dark", "doc_font", "doc_size"):
            self.s.remove(k)
        self.font.blockSignals(True)
        self.size.blockSignals(True)
        global DEFAULT_FONT, DEFAULT_SIZE
        DEFAULT_FONT, DEFAULT_SIZE = "Segoe UI", 12
        self.font.setCurrentFont(QFont(DEFAULT_FONT))
        self.size.setCurrentText("12")
        self.font.blockSignals(False)
        self.size.blockSignals(False)
        self.set_mode("system")
        self.accept()
        QTimer.singleShot(0, self.win.open_settings)  # открыть заново с исходными значениями


# ───────────────────────────── Стикеры поверх документа ─────────────────────────────
STICKY_COLORS = ["#FFE97A", "#FFC2D4", "#BDF2C9", "#BFE3FF", "#DCCBFF", "#FFD3A8", "#FFFFFF", "#2B2F3A"]


def sidecar_path(doc_path):
    p = Path(doc_path)
    return p.with_name("." + p.name + ".folio")


def read_sidecar(doc_path):
    try:
        data = json.loads(sidecar_path(doc_path).read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("stickers"), list):
            return data
    except Exception:
        pass
    return {"version": 1, "stickers": []}


_DEADLINE_CACHE = {}


def nearest_deadline(doc_path):
    sc = sidecar_path(doc_path)
    try:
        mtime = sc.stat().st_mtime
    except OSError:
        return None
    key = (str(sc), mtime)
    if key not in _DEADLINE_CACHE:
        best = None
        for st in read_sidecar(doc_path)["stickers"]:
            if st.get("deadline") and not st.get("done"):
                try:
                    d = datetime.fromisoformat(st["deadline"])
                except ValueError:
                    continue
                best = d if best is None or d < best else best
        _DEADLINE_CACHE[key] = best
    return _DEADLINE_CACHE[key]


def deadline_text(d):
    now = datetime.now()
    day = "сегодня" if d.date() == now.date() else f"{d.day} {MONTHS[d.month - 1]}"
    return f"{day} {d:%H:%M}" if (d.hour, d.minute) != (23, 59) else day


def deadline_color(d, done=False):
    if done:
        return "#30A46C"
    left = (d - datetime.now()).total_seconds()
    if left < 0:
        return "#E5484D"
    if left < 24 * 3600:
        return "#F76B15"
    return None


def move_with_sidecar(src, dst):
    sc = sidecar_path(src)
    if sc.exists() and not Path(dst).is_dir():
        try:
            shutil.move(str(sc), str(sidecar_path(dst)))
        except Exception:
            pass


class DeadlineDialog(BaseDialog):
    def __init__(self, parent, value=None):
        super().__init__(parent, "Дедлайн", "Готово")
        from PySide6.QtWidgets import QDateTimeEdit
        l = QLabel("Когда")
        l.setObjectName("Muted")
        self.lay.addWidget(l)
        self.edit = QDateTimeEdit()
        self.edit.setCalendarPopup(True)
        self.edit.setDisplayFormat("d MMMM yyyy, HH:mm")
        self.edit.setMinimumHeight(34)
        if value:
            self.edit.setDateTime(QDateTime.fromString(value, Qt.DateFormat.ISODate))
        else:
            d = QDateTime.currentDateTime().addDays(1)
            d.setTime(d.time().fromString("18:00", "HH:mm"))
            self.edit.setDateTime(d)
        self.lay.addWidget(self.edit)
        self.removed = False
        if value:
            rm = QPushButton("Убрать дедлайн")
            rm.setObjectName("Ghost")
            rm.clicked.connect(self._remove)
            self.lay.addWidget(rm, 0, Qt.AlignmentFlag.AlignLeft)
        self.finish()

    def _remove(self):
        self.removed = True
        self.accept()

    def value(self):
        return self.edit.dateTime().toString(Qt.DateFormat.ISODate)


class StickerMenu(QMenu):
    """Цвет + прозрачность стикера."""
    def __init__(self, sticker, with_colors=True):
        super().__init__(sticker)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.FramelessWindowHint | Qt.WindowType.NoDropShadowWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PySide6.QtWidgets import QSlider
        box = QWidget()
        box.setObjectName("SwatchBox")
        v = QVBoxLayout(box)
        v.setContentsMargins(12, 10, 12, 10)
        v.setSpacing(8)
        if with_colors:
            grid = QGridLayout()
            grid.setSpacing(8)
            for i, col in enumerate(STICKY_COLORS):
                b = QToolButton()
                b.setFixedSize(26, 26)
                b.setCursor(Qt.CursorShape.PointingHandCursor)
                sel = col.lower() == sticker.d.get("color", "").lower()
                b.setStyleSheet(f"QToolButton{{background:{col};border-radius:13px;"
                                f"border:{3 if sel else 1}px solid {'#3D5AFE' if sel else 'rgba(0,0,0,0.15)'};}}")
                b.clicked.connect(lambda _=False, c=col: (sticker.set_color(c), self.close()))
                grid.addWidget(b, i // 4, i % 4)
            v.addLayout(grid)
        lab = QLabel("Прозрачность")
        lab.setObjectName("Muted")
        v.addWidget(lab)
        sl = QSlider(Qt.Orientation.Horizontal)
        sl.setRange(15, 100)
        sl.setValue(round(sticker.d.get("opacity", 1) * 100))
        sl.setMinimumWidth(150)
        sl.valueChanged.connect(lambda val: sticker.set_opacity(val / 100))
        v.addWidget(sl)
        act = QWidgetAction(self)
        act.setDefaultWidget(box)
        self.addAction(act)


class Sticker(QFrame):
    HEADER = 30
    GRIP = 16
    MARGIN = 6  # место под тень

    def __init__(self, layer, data):
        super().__init__(layer.viewport)
        self.layer = layer
        self.d = data
        if not data.get("id"):
            import uuid
            data["id"] = uuid.uuid4().hex[:10]
        self._drag = None
        self._resize = None
        self._moved = False
        self.hovered = False
        self.setMouseTracking(True)
        self.setAttribute(Qt.WidgetAttribute.WA_Hover)
        self.setMinimumSize(110, 70)
        self.tools = QWidget(self)
        self.tools.setObjectName("StickerTools")
        tl = QHBoxLayout(self.tools)
        tl.setContentsMargins(0, 0, 0, 0)
        tl.setSpacing(0)
        for icon, tip, slot in self.tool_buttons():
            b = QToolButton()
            b.setFixedSize(24, 24)
            b.setIconSize(QSize(15, 15))
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setToolTip(tip)
            b.setProperty("icon_name", icon)
            b.clicked.connect(slot)
            tl.addWidget(b)
        self.tools.adjustSize()
        self.tools.hide()
        self.chip = QPushButton(self)
        self.chip.setObjectName("DeadlineChip")
        self.chip.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chip.setToolTip("Клик — отметить выполненным")
        self.chip.clicked.connect(self.toggle_done)
        self.chip.hide()

    # ── переопределяется в наследниках ──
    def tool_buttons(self):
        return [("arrow", "Стрелка от этого стикера", self.start_arrow),
                ("drop", "Цвет и прозрачность", self.color_menu), ("trash", "Удалить стикер", self.delete)]

    def bg(self):
        c = QColor(self.d.get("color", STICKY_COLORS[0]))
        c.setAlphaF(max(0.0, min(1.0, self.d.get("opacity", 0.95))))
        return c

    def fg(self):
        return "#1A1D24" if luminance(self.d.get("color", "#FFE97A")) > 0.35 else "#F2F4F8"

    def content_rect(self):
        m = self.MARGIN
        r = QRect(m + 12, m + self.HEADER, self.width() - 2 * m - 24, self.height() - 2 * m - self.HEADER - 12)
        if self.d.get("shape") == "ellipse":  # в овале текст держим ближе к центру
            dx, dy = int(self.width() * 0.11), int(self.height() * 0.08)
            r.adjust(dx, dy - 6, -dx, -dy)
        elif self.d.get("shape") == "pill":
            dx = int(min(self.width(), self.height()) * 0.18)
            r.adjust(dx, 0, -dx, 0)
        return r

    def shape_path(self, r):
        path = QPainterPath()
        shape = self.d.get("shape", "round")
        if shape == "rect":
            path.addRect(r)
        elif shape == "ellipse":
            path.addEllipse(r)
        elif shape == "pill":
            rr = min(r.width(), r.height()) / 2
            path.addRoundedRect(r, rr, rr)
        else:
            rr = min(float(self.d.get("radius", 14)), min(r.width(), r.height()) / 2)
            path.addRoundedRect(r, rr, rr)
        return path

    def is_selected(self):
        return self.layer.selected is self

    # ── отрисовка ──
    def refresh_style(self):
        fg = self.fg()
        for b in self.tools.findChildren(QToolButton):
            b.setIcon(QIcon(svg_pixmap(SVG[b.property("icon_name")], fg, size=15)))
            b.setStyleSheet("QToolButton{background:transparent;border:none;border-radius:12px;}"
                            "QToolButton:hover{background:rgba(127,127,127,0.25);}")
        self.update_chip()
        self.update()

    def update_chip(self):
        dl = self.d.get("deadline")
        if not dl:
            self.chip.hide()
            return
        try:
            d = datetime.fromisoformat(dl)
        except ValueError:
            self.chip.hide()
            return
        done = self.d.get("done", False)
        col = deadline_color(d, done)
        fg = self.fg()
        if col:
            style = f"background:{col};color:white;"
        else:
            style = f"background:rgba(127,127,127,0.18);color:{fg};"
        self.chip.setText(("✓ " if done else "⏰ ") + deadline_text(d))
        self.chip.setStyleSheet(f"QPushButton#DeadlineChip{{{style}border:none;border-radius:10px;"
                                f"padding:2px 9px;font-size:8.5pt;font-weight:600;"
                                f"{'text-decoration:line-through;' if done else ''}}}")
        self.chip.adjustSize()
        self.chip.move(self.MARGIN + 10, self.MARGIN + 6)
        self.chip.show()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        m = self.MARGIN
        r = QRectF(self.rect()).adjusted(m, m, -m, -m)
        p.setPen(Qt.PenStyle.NoPen)
        alpha = self.d.get("opacity", 1.0)
        inner = self.shape_path(r)
        if self.d.get("shadow", True):
            for i in range(m, 0, -1):  # мягкая тень — только снаружи, чтобы не просвечивала сквозь прозрачный стикер
                ring = self.shape_path(r.adjusted(-i + 1, -i + 3, i - 1, i + 1))
                p.fillPath(ring.subtracted(inner), QColor(0, 0, 0, int((7 - i) * 3 * alpha)))
        self.paint_body(p, r)
        sel = self.is_selected()
        drop = getattr(self, "drop_target", False)
        bw = float(self.d.get("border_w", 1))
        if sel or drop:
            pen = QPen(QColor(self.layer.win.th["accent"]), 3 if drop else 2)
        elif bw > 0:
            bc = QColor(self.d["border_c"]) if self.d.get("border_c") else QColor(0, 0, 0, 35)
            pen = QPen(bc, bw)
        else:
            pen = None
        if pen is not None:
            p.setPen(pen)
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawPath(inner)
        if self.hovered:  # уголок для изменения размера
            p.setPen(QPen(QColor(self.fg()), 1.4))
            x, y = r.right() - 5, r.bottom() - 5
            for k in (4, 8):
                p.drawLine(QPointF(x - k, y), QPointF(x, y - k))

    def paint_body(self, p, r):
        p.fillPath(self.shape_path(r), self.bg())

    # ── мышь: перетаскивание и размер ──
    def in_grip(self, pos):
        m = self.MARGIN
        return pos.x() >= self.width() - m - self.GRIP and pos.y() >= self.height() - m - self.GRIP

    def draggable_at(self, pos):
        return pos.y() <= self.MARGIN + self.HEADER

    def mousePressEvent(self, e):
        self.raise_()
        self._moved = False
        self.layer.select(self)
        if e.button() != Qt.MouseButton.LeftButton:
            return
        pos = e.position().toPoint()
        g = e.globalPosition().toPoint()
        if self.in_grip(pos):
            self._resize = (g, self.d["w"], self.d["h"])
        elif self.draggable_at(pos):
            self._drag = (g, self.d["x"], self.d["y"])

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        g = e.globalPosition().toPoint()
        if self._drag:
            g0, x0, y0 = self._drag
            dx, dy = g.x() - g0.x(), g.y() - g0.y()
            if abs(dx) + abs(dy) > 3:
                self._moved = True
            vw = self.layer.viewport.width()
            x, y = x0 + dx, y0 + dy
            if self.layer.snap():
                x, y = self.layer.snap_xy(self, x, y)
            self.d["x"] = max(-self.MARGIN, min(vw - 60, x))
            self.d["y"] = max(0, y)
            self.layer.place(self)
        elif self._resize:
            g0, w0, h0 = self._resize
            w = max(self.minimumWidth(), w0 + g.x() - g0.x())
            h = max(self.minimumHeight(), h0 + g.y() - g0.y())
            if self.layer.snap():
                w, h = round(w / 10) * 10, round(h / 10) * 10
            if self.d.get("type") == "photo" and self.photo_ratio():
                h = round((w - 2 * self.MARGIN) / self.photo_ratio()) + 2 * self.MARGIN
            self.d["w"], self.d["h"] = int(w), int(h)
            self._moved = True
            self.layer.place(self)
        else:
            if self.in_grip(pos):
                self.setCursor(Qt.CursorShape.SizeFDiagCursor)
            elif self.draggable_at(pos):
                self.setCursor(Qt.CursorShape.SizeAllCursor)
            else:
                self.setCursor(Qt.CursorShape.ArrowCursor)

    def mouseReleaseEvent(self, e):
        if self._drag or self._resize:
            self._drag = self._resize = None
            if self._moved:
                self.layer.mark()
                self.layer.win.inspector.refresh_values()

    def photo_ratio(self):
        return None

    def set(self, key, val):
        if val is None:
            self.d.pop(key, None)
        else:
            self.d[key] = val
        self.refresh_style()
        self.layout_content()
        self.layer.place(self)
        self.layer.mark()

    def enterEvent(self, e):
        self.hovered = True
        self.tools.show()
        self.tools.raise_()
        self.update()

    def leaveEvent(self, e):
        self.hovered = False
        self.tools.hide()
        self.update()

    def resizeEvent(self, e):
        m = self.MARGIN
        self.tools.move(self.width() - m - self.tools.width() - 6, m + 3)
        self.layout_content()
        self.update_chip()

    def layout_content(self):
        pass

    # ── действия ──
    def color_menu(self):
        menu = StickerMenu(self, with_colors=self.d.get("type") != "photo")
        btn = self.sender()
        menu.exec(btn.mapToGlobal(btn.rect().bottomLeft()) if btn else QCursor.pos())

    def set_color(self, c):
        self.d["color"] = c
        self.refresh_style()
        self.layer.mark()

    def set_opacity(self, a):
        self.d["opacity"] = a
        self.update()
        self.layer.mark()

    def edit_deadline(self):
        dlg = DeadlineDialog(self.layer.win, self.d.get("deadline"))
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return False
        if dlg.removed:
            self.d.pop("deadline", None)
            self.d["done"] = False
        else:
            self.d["deadline"] = dlg.value()
            self.d["done"] = False
        self.update_chip()
        self.layout_content()
        self.layer.mark()
        return True

    def toggle_done(self):
        self.d["done"] = not self.d.get("done", False)
        self.update_chip()
        self.layer.mark()

    def delete(self):
        self.layer.remove(self)

    def start_arrow(self):
        self.layer.add_arrow(from_sticker=self)

    def doc_rect(self):
        m = self.MARGIN
        return QRectF(self.d["x"] + m, self.d["y"] + m, self.d["w"] - 2 * m, self.d["h"] - 2 * m)


def border_point(rect, toward, gap=7):
    """Точка на краю прямоугольника по направлению к toward (плюс небольшой зазор)."""
    c = rect.center()
    dx, dy = toward.x() - c.x(), toward.y() - c.y()
    if abs(dx) < 1e-6 and abs(dy) < 1e-6:
        return c
    tx = (rect.width() / 2) / abs(dx) if abs(dx) > 1e-6 else float("inf")
    ty = (rect.height() / 2) / abs(dy) if abs(dy) > 1e-6 else float("inf")
    t = min(tx, ty)
    length = (dx * dx + dy * dy) ** 0.5
    t += gap / length
    return QPointF(c.x() + dx * t, c.y() + dy * t)


ARROW_COLORS = ["#1A1D24", "#5F6368", "#E5484D", "#F76B15", "#F5B400", "#30A46C",
                "#12A594", "#0090FF", "#3D5AFE", "#8E4EC6", "#D6409F", "#FFFFFF"]


class ArrowSticker(QWidget):
    """Стрелка: концы тянутся, середина изгибает линию, концы цепляются к стикерам."""
    PAD = 26
    HANDLE = 7

    def __init__(self, layer, data):
        super().__init__(layer.viewport)
        self.layer = layer
        self.d = data
        if not data.get("id"):
            import uuid
            data["id"] = uuid.uuid4().hex[:10]
        data.setdefault("bend", [0, 0])
        if "heads" in data and "head_end" not in data:  # старый формат
            h = data.pop("heads")
            data["head_end"] = "none" if h == "none" else "arrow"
            data["head_start"] = "arrow" if h == "both" else "none"
        if "dash" in data and isinstance(data["dash"], bool):
            data["line"] = "dash" if data.pop("dash") else "solid"
        self.hover = False
        self._drag = None
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.menu)
        self.setAttribute(Qt.WidgetAttribute.WA_Hover)

    # ── геометрия в координатах документа ──
    def ends(self):
        s1 = self.layer.by_id(self.d.get("a1"))
        s2 = self.layer.by_id(self.d.get("a2"))
        r1 = s1.doc_rect().center() if s1 else QPointF(self.d["x1"], self.d["y1"])
        r2 = s2.doc_rect().center() if s2 else QPointF(self.d["x2"], self.d["y2"])
        bx, by = self.d.get("bend", [0, 0])
        h = QPointF((r1.x() + r2.x()) / 2 + bx, (r1.y() + r2.y()) / 2 + by)
        p1 = border_point(s1.doc_rect(), h) if s1 else r1
        p2 = border_point(s2.doc_rect(), h) if s2 else r2
        h = QPointF((p1.x() + p2.x()) / 2 + bx, (p1.y() + p2.y()) / 2 + by)
        c = QPointF(2 * h.x() - (p1.x() + p2.x()) / 2, 2 * h.y() - (p1.y() + p2.y()) / 2)
        return p1, p2, h, c

    @property
    def selected(self):
        return self.layer.selected is self

    def kind(self):
        return self.d.get("shape", "curve")

    def elbow_points(self):
        p1, p2, h, c = self.ends()
        mx = (p1.x() + p2.x()) / 2 + self.d.get("bend", [0, 0])[0]
        return [p1, QPointF(mx, p1.y()), QPointF(mx, p2.y()), p2]

    def handle_point(self):
        p1, p2, h, c = self.ends()
        if self.kind() == "elbow":
            pts = self.elbow_points()
            return QPointF(pts[1].x(), (p1.y() + p2.y()) / 2)
        return h

    def doc_path(self):
        p1, p2, h, c = self.ends()
        path = QPainterPath(p1)
        k = self.kind()
        if k == "straight":
            path.lineTo(p2)
        elif k == "elbow":
            for pt in self.elbow_points()[1:]:
                path.lineTo(pt)
        else:
            path.quadTo(c, p2)
        return path

    def tangents(self):
        """Точки, откуда «приходит» линия к концам — для наконечников."""
        p1, p2, h, c = self.ends()
        k = self.kind()
        if k == "straight":
            return p2, p1
        if k == "elbow":
            pts = self.elbow_points()
            a = pts[1] if (pts[1] - p1).manhattanLength() > 1 else pts[2]
            b = pts[2] if (pts[2] - p2).manhattanLength() > 1 else pts[1]
            return a, b
        return c, c

    def to_doc(self, global_pos):
        v = self.layer.viewport.mapFromGlobal(global_pos)
        return QPointF(v.x(), v.y() + self.layer.scroll())

    def origin(self):
        return QPointF(self.x(), self.y() + self.layer.scroll())

    def update_geometry(self):
        p1, p2, h, c = self.ends()
        path = self.doc_path()
        h = self.handle_point()
        box = path.controlPointRect().united(QRectF(h, QSizeF(1, 1))).adjusted(-self.PAD, -self.PAD, self.PAD, self.PAD)
        self.setGeometry(int(box.x()), int(box.y() - self.layer.scroll()), int(box.width()) + 1, int(box.height()) + 1)
        # маска: кликаются только сама линия и ручки — текст под стрелкой остаётся доступным
        from PySide6.QtGui import QPainterPathStroker, QRegion
        o = self.origin()
        local = QPainterPath(path)
        local.translate(-o.x(), -o.y())
        st = QPainterPathStroker()
        st.setWidth(max(16, float(self.d.get("width", 2.5)) + 14))
        region = QRegion(st.createStroke(local).toFillPolygon().toPolygon())
        for pt in (p1, p2):
            region = region.united(QRegion(int(pt.x() - o.x() - 22), int(pt.y() - o.y() - 22), 44, 44,
                                           QRegion.RegionType.Ellipse))
        if self.selected and self.kind() != "straight":
            region = region.united(QRegion(int(h.x() - o.x() - 12), int(h.y() - o.y() - 12), 24, 24,
                                           QRegion.RegionType.Ellipse))
        self.setMask(region)
        self.update()

    # ── рисование ──
    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        o = self.origin()
        p.translate(-o.x(), -o.y())
        p1, p2, _, _ = self.ends()
        h = self.handle_point()
        col = QColor(self.d.get("color", "#3D5AFE"))
        w = float(self.d.get("width", 2.5))
        path = self.doc_path()
        if self.selected or self.hover:  # мягкое свечение под линией
            glow = QColor(self.layer.win.th["accent"])
            glow.setAlpha(45)
            p.setPen(QPen(glow, w + 8, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            p.drawPath(path)
        p.setOpacity(float(self.d.get("opacity", 1.0)))
        style = {"dash": Qt.PenStyle.DashLine, "dot": Qt.PenStyle.DotLine}.get(self.d.get("line"), Qt.PenStyle.SolidLine)
        p.setPen(QPen(col, w, style, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawPath(path)
        t1, t2 = self.tangents()
        self.draw_head(p, p2, t2, w, col, self.d.get("head_end", "arrow"))
        self.draw_head(p, p1, t1, w, col, self.d.get("head_start", "none"))
        p.setOpacity(1)
        if self.selected:
            acc = QColor(self.layer.win.th["accent"])
            pts = [(p1, self.HANDLE, False), (p2, self.HANDLE, False)]
            if self.kind() != "straight":
                pts.append((h, self.HANDLE - 1, True))
            for pt, r, filled in pts:
                p.setPen(QPen(acc, 2))
                p.setBrush(acc if filled else QColor("#FFFFFF"))
                p.drawEllipse(pt, r, r)

    @staticmethod
    def draw_head(p, tip, ctrl, w, col, kind):
        import math
        from PySide6.QtGui import QPolygonF
        if kind == "none":
            return
        ang = math.atan2(tip.y() - ctrl.y(), tip.x() - ctrl.x())
        size = 9 + w * 2.2
        p.setPen(QPen(col, w, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        def at(a, r):
            return QPointF(tip.x() + r * math.cos(a), tip.y() + r * math.sin(a))
        if kind == "arrow":
            p.setBrush(col)
            p.drawPolygon(QPolygonF([tip, at(ang + math.radians(150), size), at(ang - math.radians(150), size)]))
        elif kind == "open":
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawLine(tip, at(ang + math.radians(145), size))
            p.drawLine(tip, at(ang - math.radians(145), size))
        elif kind == "circle":
            p.setBrush(col)
            r = 3 + w * 1.2
            p.drawEllipse(at(ang + math.pi, r), r, r)
        elif kind == "diamond":
            p.setBrush(col)
            r = 5 + w * 1.6
            back = at(ang + math.pi, 2 * r)
            mid = at(ang + math.pi, r)
            side = QPointF(math.cos(ang + math.pi / 2) * r * 0.7, math.sin(ang + math.pi / 2) * r * 0.7)
            p.drawPolygon(QPolygonF([tip, mid + side, back, mid - side]))
        elif kind == "bar":
            r = 5 + w * 1.6
            side = QPointF(math.cos(ang + math.pi / 2) * r, math.sin(ang + math.pi / 2) * r)
            p.drawLine(tip + side, tip - side)

    # ── мышь ──
    def hit(self, doc_pt):
        p1, p2, h, c = self.ends()
        def near(a, r=11):
            return (a.x() - doc_pt.x()) ** 2 + (a.y() - doc_pt.y()) ** 2 <= r * r
        if near(p1):
            return "p1"
        if near(p2):
            return "p2"
        if self.selected and self.kind() != "straight" and near(self.handle_point()):
            return "h"
        return "line"

    def mousePressEvent(self, e):
        if e.button() != Qt.MouseButton.LeftButton:
            return
        self.setFocus()
        self.layer.select(self)
        doc = self.to_doc(e.globalPosition().toPoint())
        part = self.hit(doc)
        p1, p2, h, c = self.ends()
        self._drag = (part, doc, p1, p2, list(self.d.get("bend", [0, 0])))

    def mouseMoveEvent(self, e):
        doc = self.to_doc(e.globalPosition().toPoint())
        if not self._drag:
            part = self.hit(doc)
            self.setCursor(Qt.CursorShape.SizeAllCursor if part in ("p1", "p2", "h", "line")
                           else Qt.CursorShape.ArrowCursor)
            return
        part, start, p1, p2, bend = self._drag
        dx, dy = doc.x() - start.x(), doc.y() - start.y()
        if part == "p1":
            self.d["a1"] = None
            self.d["x1"], self.d["y1"] = doc.x(), doc.y()
            self.layer.highlight_drop(doc, exclude=self.d.get("a2"))
        elif part == "p2":
            self.d["a2"] = None
            self.d["x2"], self.d["y2"] = doc.x(), doc.y()
            self.layer.highlight_drop(doc, exclude=self.d.get("a1"))
        elif part == "h":
            self.d["bend"] = [bend[0] + dx, bend[1] + (dy if self.kind() != "elbow" else 0)]
        else:  # двигаем всю стрелку (свободные концы)
            if not self.d.get("a1"):
                self.d["x1"], self.d["y1"] = p1.x() + dx, p1.y() + dy
            if not self.d.get("a2"):
                self.d["x2"], self.d["y2"] = p2.x() + dx, p2.y() + dy
        self.update_geometry()

    def mouseReleaseEvent(self, e):
        if not self._drag:
            return
        part = self._drag[0]
        self._drag = None
        doc = self.to_doc(e.globalPosition().toPoint())
        if part in ("p1", "p2"):  # отпустили конец на стикере — прицепляем
            key, other = ("a1", "a2") if part == "p1" else ("a2", "a1")
            target = self.layer.sticker_at(doc, exclude=self.d.get(other))
            if target is not None:
                self.d[key] = target.d["id"]
            self.layer.highlight_drop(None)
        self.update_geometry()
        self.layer.mark()
        self.layer.win.inspector.refresh_values()

    def mouseDoubleClickEvent(self, e):
        if self.hit(self.to_doc(e.globalPosition().toPoint())) == "h":
            self.straighten()

    def enterEvent(self, e):
        self.hover = True
        self.update()

    def leaveEvent(self, e):
        self.hover = False
        self.update()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self.delete()
            return
        super().keyPressEvent(e)

    def select(self, on):
        if on:
            self.layer.select(self)
        self.update_geometry()

    # ── меню ──
    def menu(self, pos):
        self.layer.select(self)
        m = styled_menu(self)
        m.addAction("Настроить…  (панель справа)", lambda: self.layer.win.inspector.show_for(self))
        m.addAction("Выпрямить", self.straighten)
        m.addAction("Развернуть", self.reverse)
        if self.d.get("a1") or self.d.get("a2"):
            m.addAction("Отцепить от стикеров", self.detach)
        m.addSeparator()
        m.addAction("Удалить стрелку", self.delete)
        m.exec(self.mapToGlobal(pos))

    def set(self, key, val):
        if val is None:
            self.d.pop(key, None)
        else:
            self.d[key] = val
        self.update_geometry()
        self.layer.mark()

    def straighten(self):
        self.set("bend", [0, 0])

    def reverse(self):
        p1, p2, h, c = self.ends()
        d = self.d
        d["a1"], d["a2"] = d.get("a2"), d.get("a1")
        d["x1"], d["y1"], d["x2"], d["y2"] = p2.x(), p2.y(), p1.x(), p1.y()
        self.update_geometry()
        self.layer.mark()

    def detach(self):
        self.freeze()
        self.d["a1"] = self.d["a2"] = None
        self.update_geometry()
        self.layer.mark()

    def freeze(self, only=None):
        """Запомнить текущие координаты концов (перед тем как отцепить)."""
        p1, p2, h, c = self.ends()
        if only in (None, self.d.get("a1")):
            self.d["x1"], self.d["y1"] = p1.x(), p1.y()
        if only in (None, self.d.get("a2")):
            self.d["x2"], self.d["y2"] = p2.x(), p2.y()

    def refresh_style(self):
        self.update()

    def delete(self):
        self.layer.remove(self)


class NoteEdit(SmartEdit):
    def focusInEvent(self, e):
        super().focusInEvent(e)
        st = self.parent()
        if isinstance(st, Sticker):
            st.layer.select(st)


class NoteSticker(Sticker):
    def __init__(self, layer, data):
        super().__init__(layer, data)
        self.body = NoteEdit(self)
        self.body.setFrameShape(QFrame.Shape.NoFrame)
        self.body.setObjectName("NoteBody")
        self.body.viewport().setAutoFillBackground(False)
        self.body.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.body.setPlaceholderText("Заметка…")
        f = QFont(DEFAULT_FONT)
        f.setPointSizeF(11)
        self.body.document().setDefaultFont(f)
        self.body.document().setDocumentMargin(2)
        if data.get("html"):
            self.body.setHtml(data["html"])
        self.body.textChanged.connect(self._text_changed)
        self.body.cursorPositionChanged.connect(lambda: layer.win.sync_toolbar() if layer.win.tgt() is self.body else None)
        self.body.currentCharFormatChanged.connect(lambda _: layer.win.sync_toolbar() if layer.win.tgt() is self.body else None)
        self.refresh_style()

    def tool_buttons(self):
        return [("clock", "Дедлайн", self.edit_deadline), ("checklist", "Список дел", self.add_checklist),
                ("arrow", "Стрелка от этой заметки", self.start_arrow),
                ("drop", "Цвет и прозрачность", self.color_menu), ("trash", "Удалить заметку", self.delete)]

    def fg(self):
        return self.d.get("text_c") or super().fg()

    def refresh_style(self):
        fg = self.fg()
        self.body.setStyleSheet(f"QTextEdit#NoteBody{{background:transparent;border:none;color:{fg};"
                                f"font-family:'{DEFAULT_FONT}';font-size:11pt;"
                                f"selection-background-color:rgba(61,90,254,0.35);selection-color:{fg};}}")
        super().refresh_style()

    def layout_content(self):
        r = self.content_rect()
        top = r.top()
        if self.chip.isVisible() or self.d.get("deadline"):
            top = self.MARGIN + self.HEADER + 2
        self.body.setGeometry(r.left() - 4, top, r.width() + 8, r.bottom() - top + 6)

    def draggable_at(self, pos):
        return pos.y() <= self.MARGIN + self.HEADER and not self.chip.geometry().contains(pos)

    def _text_changed(self):
        self.d["html"] = self.body.toHtml()
        self.layer.mark()

    def add_checklist(self):
        self.body.setFocus()
        c = self.body.textCursor()
        if c.currentList() is None:
            if c.block().text().strip():
                c.movePosition(QTextCursor.MoveOperation.EndOfBlock)
                c.insertBlock()
            make_list(c, "check")
            self.body.setTextCursor(c)


class PhotoSticker(Sticker):
    def __init__(self, layer, data):
        super().__init__(layer, data)
        self.img = QImage()
        try:
            import base64
            self.img = QImage.fromData(base64.b64decode(data.get("image", "")))
        except Exception:
            pass
        self.setMinimumSize(70, 60)
        self.refresh_style()

    def tool_buttons(self):
        return [("crop", "Обрезать", self.crop), ("arrow", "Стрелка от этого фото", self.start_arrow),
                ("drop", "Прозрачность", self.color_menu), ("trash", "Удалить фото", self.delete)]

    def crop_reset(self):
        ratio = self.photo_ratio()
        if ratio:
            self.d["h"] = round((self.d["w"] - 2 * self.MARGIN) / ratio) + 2 * self.MARGIN
            self.layer.place(self)
            self.update()

    def crop(self):
        if self.img.isNull():
            return
        dlg = CropDialog(self.layer.win, self.img, self.d.get("crop", [0, 0, 1, 1]))
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.d["crop"] = dlg.result_crop()
            ratio = self.photo_ratio()
            self.d["h"] = round((self.d["w"] - 2 * self.MARGIN) / ratio) + 2 * self.MARGIN
            self.layer.place(self)
            self.update()
            self.layer.mark()

    def fg(self):
        return "#FFFFFF"

    def draggable_at(self, pos):
        return True

    def crop_rect(self):
        if self.img.isNull():
            return QRectF()
        cx, cy, cw, ch = self.d.get("crop", [0, 0, 1, 1])
        return QRectF(cx * self.img.width(), cy * self.img.height(), cw * self.img.width(), ch * self.img.height())

    def photo_ratio(self):
        cr = self.crop_rect()
        return cr.width() / max(1, cr.height()) if not cr.isEmpty() else self.d.get("ratio")

    def paint_body(self, p, r):
        path = self.shape_path(r)
        p.save()
        p.setClipPath(path)
        p.setOpacity(self.d.get("opacity", 1.0))
        if not self.img.isNull():
            src = self.crop_rect()
            # заполняем форму без искажения (как «cover»)
            tr = r.width() / max(1, r.height())
            sr = src.width() / max(1, src.height())
            if sr > tr:
                nw = src.height() * tr
                src = QRectF(src.x() + (src.width() - nw) / 2, src.y(), nw, src.height())
            else:
                nh = src.width() / tr
                src = QRectF(src.x(), src.y() + (src.height() - nh) / 2, src.width(), nh)
            p.drawImage(r, self.img, src)
        else:
            p.fillPath(path, QColor("#888"))
        if self.hovered:  # затемнение сверху, чтобы кнопки читались
            p.setOpacity(1)
            from PySide6.QtGui import QLinearGradient
            g = QLinearGradient(0, r.top(), 0, r.top() + 44)
            g.setColorAt(0, QColor(0, 0, 0, 110))
            g.setColorAt(1, QColor(0, 0, 0, 0))
            p.fillRect(r, g)
        p.restore()


class LinkSticker(Sticker):
    def __init__(self, layer, data):
        super().__init__(layer, data)
        self.setMinimumSize(170, 84)
        self.refresh_style()

    def tool_buttons(self):
        return [("pencil", "Изменить ссылку", self.edit_link), ("arrow", "Стрелка от этой ссылки", self.start_arrow),
                ("drop", "Цвет и прозрачность", self.color_menu), ("trash", "Удалить", self.delete)]

    def draggable_at(self, pos):
        return True

    def paint_body(self, p, r):
        super().paint_body(p, r)
        fg = QColor(self.fg())
        url = QUrl(self.d.get("url", ""))
        icon = cached_svg("file" if url.isLocalFile() else "globe", self.layer.win.th["accent"], 30)
        circle = QRectF(r.left() + 14, r.center().y() - 22, 44, 44)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(255, 255, 255, 170) if luminance(self.d.get("color", "#FFF")) > 0.35 else QColor(255, 255, 255, 40))
        p.drawEllipse(circle)
        p.drawPixmap(int(circle.center().x() - 15), int(circle.center().y() - 15), icon)
        tx = circle.right() + 12
        f = QFont(self.layer.win.ui_font)
        f.setPointSizeF(10.5)
        f.setWeight(QFont.Weight.DemiBold)
        p.setFont(f)
        p.setPen(fg)
        fm = QFontMetrics(f)
        w = int(r.right() - tx - 12)
        p.drawText(QRectF(tx, r.center().y() - fm.height() + 1, w, fm.height()), 0,
                   fm.elidedText(self.d.get("title") or "Ссылка", Qt.TextElideMode.ElideRight, w))
        f2 = QFont(self.layer.win.ui_font)
        f2.setPointSizeF(8.5)
        p.setFont(f2)
        c = QColor(fg)
        c.setAlpha(150)
        p.setPen(c)
        label = Path(url.toLocalFile()).name if url.isLocalFile() else url.host().removeprefix("www.")
        p.drawText(QRectF(tx, r.center().y() + 3, w, 18), 0,
                   QFontMetrics(f2).elidedText(label, Qt.TextElideMode.ElideMiddle, w))

    def mouseReleaseEvent(self, e):
        was_drag = self._moved
        super().mouseReleaseEvent(e)
        pos = e.position().toPoint()
        if not was_drag and e.button() == Qt.MouseButton.LeftButton and not self.in_grip(pos) \
                and not self.tools.geometry().contains(pos):
            open_href(self.d.get("url", ""), self.layer.win)

    def edit_link(self):
        dlg = LinkDialog(self.layer.win, "Изменить ссылку", "Сохранить", "Название",
                         name=self.d.get("title", ""), url=self.d.get("url", ""))
        if dlg.exec() == QDialog.DialogCode.Accepted:
            name, url = dlg.values()
            if url:
                self.d["title"], self.d["url"] = name, url
                self.update()
                self.layer.mark()


class StickerLayer(QObject):
    """Стикеры живут поверх текста, прокручиваются вместе с ним и хранятся в скрытом файле рядом с документом."""
    def __init__(self, win, editor):
        super().__init__(win)
        self.win = win
        self.editor = editor
        self.viewport = editor.viewport()
        self.items = []
        self.path = None
        self.selected = None
        editor.verticalScrollBar().valueChanged.connect(lambda _: self.relayout())
        self.viewport.installEventFilter(self)
        self.save_timer = QTimer(self)
        self.save_timer.setSingleShot(True)
        self.save_timer.setInterval(600)
        self.save_timer.timeout.connect(self.save)

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.Type.Resize:
            QTimer.singleShot(0, self.relayout)
        return False

    def scroll(self):
        return self.editor.verticalScrollBar().value()

    def place(self, s):
        if isinstance(s, ArrowSticker):
            s.update_geometry()
            s.lower()
            return
        vw = self.viewport.width()
        s.d["x"] = max(-Sticker.MARGIN, min(s.d["x"], vw - min(s.d["w"], vw) + Sticker.MARGIN))
        s.setGeometry(int(s.d["x"]), int(s.d["y"] - self.scroll()), int(s.d["w"]), int(s.d["h"]))
        sid = s.d.get("id")
        for a in self.items:  # стрелки, прицепленные к этому стикеру, едут вместе с ним
            if isinstance(a, ArrowSticker) and sid and sid in (a.d.get("a1"), a.d.get("a2")):
                a.update_geometry()
        self.win.glass.schedule()

    def relayout(self):
        for s in self.items:
            if not isinstance(s, ArrowSticker):
                self.place(s)
        for s in self.items:
            if isinstance(s, ArrowSticker):
                self.place(s)

    def select(self, item):
        if item is self.selected:
            if item is not None and not self.win.inspector.isVisible() and self.win.inspector.auto():
                self.win.inspector.show_for(item)
            return
        old = self.selected
        self.selected = item
        for it in (old, item):
            if it is not None:
                try:
                    if isinstance(it, ArrowSticker):
                        it.update_geometry()
                    else:
                        it.update()
                except RuntimeError:
                    pass
        if item is None:
            self.win.inspector.hide_panel()
        elif self.win.inspector.auto() or self.win.inspector.isVisible():
            self.win.inspector.show_for(item)

    def snap(self):
        return self.win.settings.value("snap", "true") in (True, "true")

    def snap_xy(self, st, x, y):
        """Притягивание к сетке 10 px и к краям соседних стикеров."""
        m = Sticker.MARGIN
        w, h = st.d["w"], st.d["h"]
        gx, gy = round((x + m) / 10) * 10 - m, round((y + m) / 10) * 10 - m
        best_x, best_y = gx, gy
        for o in self.items:
            if o is st or isinstance(o, ArrowSticker):
                continue
            for ox in (o.d["x"], o.d["x"] + o.d["w"] - w, o.d["x"] + (o.d["w"] - w) / 2):
                if abs(ox - x) < 8:
                    best_x = ox
            for oy in (o.d["y"], o.d["y"] + o.d["h"] - h, o.d["y"] + (o.d["h"] - h) / 2):
                if abs(oy - y) < 8:
                    best_y = oy
        return best_x, best_y

    def bring(self, st, front=True):
        if st not in self.items:
            return
        self.items.remove(st)
        if front:
            self.items.append(st)
            st.raise_()
        else:
            self.items.insert(0, st)
            st.lower()
            for a in self.items:
                if isinstance(a, ArrowSticker):
                    a.lower()
        self.mark()

    def by_id(self, sid):
        if not sid:
            return None
        for s in self.items:
            if not isinstance(s, ArrowSticker) and s.d.get("id") == sid:
                return s
        return None

    def sticker_at(self, doc_pt, exclude=None):
        for s in reversed(self.items):
            if isinstance(s, ArrowSticker) or s.d.get("id") == exclude:
                continue
            if s.doc_rect().contains(doc_pt):
                return s
        return None

    def highlight_drop(self, doc_pt, exclude=None):
        target = self.sticker_at(doc_pt, exclude) if doc_pt is not None else None
        for s in self.items:
            if not isinstance(s, ArrowSticker):
                flag = s is target
                if getattr(s, "drop_target", False) != flag:
                    s.drop_target = flag
                    s.update()

    def add_arrow(self, from_sticker=None):
        col = self.win.th["accent"]
        if from_sticker is not None:
            r = from_sticker.doc_rect()
            x1, y1 = r.right(), r.center().y()
            x2 = min(self.viewport.width() - 30, r.right() + 170)
            if x2 - x1 < 80:
                x2 = max(30, r.left() - 170)
            d = dict(type="arrow", a1=from_sticker.d["id"], a2=None, x1=x1, y1=y1, x2=x2, y2=y1 + 30,
                     bend=[0, -24], color=col, width=2.5, head_end="arrow", head_start="none")
        else:
            vw = self.viewport.width()
            y = self.scroll() + 260
            d = dict(type="arrow", a1=None, a2=None, x1=vw * 0.3, y1=y, x2=vw * 0.62, y2=y,
                     bend=[0, -30], color=col, width=2.5, head_end="arrow", head_start="none")
        a = self.create(d)
        a.setFocus()
        a.select(True)
        self.mark()
        return a

    def clear(self):
        self.selected = None
        if hasattr(self.win, "inspector"):
            self.win.inspector.hide_panel()
        for s in self.items:
            s.hide()
            s.setParent(None)
            s.deleteLater()
        self.items = []

    def load(self, doc_path):
        if self.save_timer.isActive():
            self.save_timer.stop()
            self.save()
        self.clear()
        self.path = doc_path
        if doc_path:
            for d in read_sidecar(doc_path)["stickers"]:
                try:
                    self.create(d)
                except Exception:
                    continue
            self.relayout()

    def create(self, d):
        cls = {"note": NoteSticker, "photo": PhotoSticker, "link": LinkSticker,
               "arrow": ArrowSticker}[d.get("type", "note")]
        s = cls(self, d)
        s.show()
        self.items.append(s)
        self.place(s)
        return s

    def mark(self):
        self.save_timer.start()
        self.win.glass.schedule()

    def save(self):
        if not self.path:
            return
        sc = sidecar_path(self.path)
        if not self.items:
            if sc.exists():
                try:
                    sc.unlink()
                except OSError:
                    pass
            return
        try:
            _write_hidden(sc, json.dumps({"version": 1, "stickers": [s.d for s in self.items]}, ensure_ascii=False))
        except OSError:
            pass

    def set_path(self, path):
        self.path = path
        self.save()

    def refresh_styles(self):
        for s in self.items:
            s.refresh_style()

    def spot(self, w, h, pos=None):
        vw = self.viewport.width()
        if pos is not None:
            return max(0, min(pos.x() - w // 2, vw - w)), max(0, pos.y() + self.scroll() - 20)
        n = len(self.items) % 5
        return max(0, vw - w - 24 - 18 * n), self.scroll() + 96 + 26 * n

    def add_note(self, kind="note"):
        w, h = 250, 210
        x, y = self.spot(w, h)
        d = dict(type="note", x=x, y=y, w=w, h=h, color=STICKY_COLORS[0], opacity=0.95, html="")
        if kind == "todo":
            d["color"] = STICKY_COLORS[2]
        s = self.create(d)
        s.body.setFocus()
        if kind == "todo":
            s.add_checklist()
        if kind == "deadline":
            d["color"] = STICKY_COLORS[1]
            s.refresh_style()
            if not s.edit_deadline():
                self.remove(s)
                return None
            s.body.setFocus()
        self.mark()
        return s

    def add_photo(self, src, pos=None):
        import base64
        if isinstance(src, QImage):
            img = src
            buf = QBuffer()
            buf.open(QIODevice.OpenModeFlag.WriteOnly)
            img.save(buf, "PNG")
            raw = bytes(buf.data())
        else:
            raw = Path(src).read_bytes()
            img = QImage.fromData(raw)
        if img.isNull():
            QMessageBox.warning(self.win, APP_NAME, "Не получилось открыть картинку.")
            return None
        # большие фото уменьшаем, чтобы файл заметок не разрастался
        if max(img.width(), img.height()) > 1600:
            img = img.scaled(1600, 1600, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            buf = QBuffer()
            buf.open(QIODevice.OpenModeFlag.WriteOnly)
            img.save(buf, "JPG", 88)
            raw = bytes(buf.data())
        ratio = img.width() / max(1, img.height())
        w = 260
        h = round((w - 2 * Sticker.MARGIN) / ratio) + 2 * Sticker.MARGIN
        x, y = self.spot(w, h, pos)
        d = dict(type="photo", x=x, y=y, w=w, h=h, ratio=ratio, opacity=1.0, image=base64.b64encode(raw).decode())
        s = self.create(d)
        s.raise_()
        self.mark()
        return s

    def add_link(self):
        clip = QGuiApplication.clipboard().text().strip()
        url = clip if re.match(r"^https?://\S+$", clip) else ""
        dlg = LinkDialog(self.win, "Ссылка-стикер", "Добавить", "Название", url=url)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        name, url = dlg.values()
        if not url:
            return None
        w, h = 250, 96
        x, y = self.spot(w, h)
        s = self.create(dict(type="link", x=x, y=y, w=w, h=h, color="#FFFFFF", opacity=0.97, title=name, url=url))
        self.mark()
        return s

    def remove(self, s):
        if not isinstance(s, ArrowSticker):  # прицепленные стрелки остаются на месте
            sid = s.d.get("id")
            for a in self.items:
                if isinstance(a, ArrowSticker) and sid in (a.d.get("a1"), a.d.get("a2")):
                    a.freeze(sid)
                    if a.d.get("a1") == sid:
                        a.d["a1"] = None
                    if a.d.get("a2") == sid:
                        a.d["a2"] = None
        if s in self.items:
            self.items.remove(s)
        if self.selected is s:
            self.select(None)
        if self.win.tgt() is getattr(s, "body", None):
            self.win.set_target(self.win.editor)
        s.hide()
        s.setParent(None)
        s.deleteLater()
        self.relayout()
        self.mark()


# ───────────────────────────── Стеклянная панель ─────────────────────────────
class GlassBar(QFrame):
    """Плавающая панель форматирования с живым размытием того, что под ней (как «стекло» в iOS)."""
    PAD = 10  # поля под тень

    def __init__(self, win, radius=None):
        super().__init__()
        self.win = win
        self.radius = radius
        self.setObjectName("GlassBar")
        self.blur = None
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(12)
        self.timer.timeout.connect(self.refresh_blur)

    def enabled(self):
        return self.win.settings.value("glass", "true") in (True, "true")

    def schedule(self):
        if self.isVisible():
            self.timer.start()

    def refresh_blur(self):
        if not self.enabled() or not self.isVisible():
            self.blur = None
            self.update()
            return
        ed = self.win.editor
        w, h = self.width(), self.height()
        canvas = QPixmap(w, h)
        canvas.fill(QColor(self.win.th["bg"]))
        top_left = ed.mapFrom(self.parentWidget(), self.pos())
        src = QRect(top_left, QSize(w, h)).intersected(ed.rect())
        if not src.isEmpty():
            grab = ed.grab(src)
            p = QPainter(canvas)
            p.drawPixmap(src.topLeft() - top_left, grab)
            p.end()
        img = canvas.toImage()
        blur = int(self.win.settings.value("glass_blur", 12))
        for k in (max(2, blur), max(2, blur // 2)):  # два прохода: уменьшить-увеличить = мягкое размытие
            small = img.scaled(max(1, w // k), max(1, h // k), Qt.AspectRatioMode.IgnoreAspectRatio,
                               Qt.TransformationMode.SmoothTransformation)
            img = small.scaled(w, h, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.blur = QPixmap.fromImage(img)
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        th = self.win.th
        dark = self.win.theme == "dark"
        pad = self.PAD
        r = QRectF(self.rect()).adjusted(pad, pad - 4, -pad, -pad)
        rf = getattr(self.win, "rf", 1.0)
        radius = (r.height() / 2 if self.radius is None else self.radius) * rf
        radius = min(radius, r.height() / 2, r.width() / 2)
        path = QPainterPath()
        path.addRoundedRect(r, radius, radius)
        # тень
        p.setPen(Qt.PenStyle.NoPen)
        for i in range(pad, 0, -2):
            p.setBrush(QColor(0, 0, 0, (60 if dark else 16) // (i // 2 + 1)))
            p.drawRoundedRect(r.adjusted(-i / 2, -i / 2 + 2, i / 2, i / 2 + 3), radius + i / 2, radius + i / 2)
        if self.enabled() and self.blur is not None:
            from PySide6.QtGui import QLinearGradient
            p.save()
            p.setClipPath(path)
            p.drawPixmap(0, 0, self.blur)
            tint = QColor(th["surface"])
            tint.setAlphaF(float(self.win.settings.value("glass_tint", 0.40)))
            p.fillPath(path, tint)
            g = QLinearGradient(0, r.top(), 0, r.bottom())
            g.setColorAt(0, QColor(255, 255, 255, 30 if dark else 120))
            g.setColorAt(0.55, QColor(255, 255, 255, 0))
            p.fillPath(path, g)
            p.restore()
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.setPen(QPen(QColor(255, 255, 255, 45 if dark else 210), 1.2))
            p.drawPath(path)
            inner = QPainterPath()
            inner.addRoundedRect(r.adjusted(1.5, 1.5, -1.5, -1.5), radius - 1.5, radius - 1.5)
            p.setPen(QPen(QColor(0, 0, 0, 40 if dark else 14), 1))
            p.drawPath(inner)
        else:
            p.setBrush(QColor(th["surface"]))
            p.setPen(QPen(QColor(th["border"]), 1))
            p.drawPath(path)


# ───────────────────────────── Обрезка фото ─────────────────────────────
class CropView(QWidget):
    def __init__(self, img, crop):
        super().__init__()
        self.img = img
        self.crop = QRectF(*crop)  # в долях 0..1
        self.ratio = None  # None = свободно
        self._drag = None
        self.setMinimumSize(520, 380)
        self.setMouseTracking(True)

    def img_rect(self):
        r = QRectF(self.rect()).adjusted(12, 12, -12, -12)
        ir = self.img.width() / max(1, self.img.height())
        if r.width() / r.height() > ir:
            w = r.height() * ir
            return QRectF(r.x() + (r.width() - w) / 2, r.y(), w, r.height())
        h = r.width() / ir
        return QRectF(r.x(), r.y() + (r.height() - h) / 2, r.width(), h)

    def crop_px(self):
        ir = self.img_rect()
        return QRectF(ir.x() + self.crop.x() * ir.width(), ir.y() + self.crop.y() * ir.height(),
                      self.crop.width() * ir.width(), self.crop.height() * ir.height())

    def set_ratio(self, ratio):
        self.ratio = ratio
        if ratio:
            ir = self.img_rect()
            c = self.crop_px()
            w, h = c.width(), c.height()
            if w / h > ratio:
                w = h * ratio
            else:
                h = w / ratio
            cx, cy = c.center().x(), c.center().y()
            self._set_px(QRectF(cx - w / 2, cy - h / 2, w, h))
        self.update()

    def _set_px(self, r):
        ir = self.img_rect()
        r = r.intersected(ir) if not ir.contains(r) else r
        self.crop = QRectF((r.x() - ir.x()) / ir.width(), (r.y() - ir.y()) / ir.height(),
                           r.width() / ir.width(), r.height() / ir.height())

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        ir = self.img_rect()
        p.drawImage(ir, self.img)
        c = self.crop_px()
        outside = QPainterPath()
        outside.addRect(ir)
        inner = QPainterPath()
        inner.addRect(c)
        p.fillPath(outside.subtracted(inner), QColor(0, 0, 0, 140))
        p.setPen(QPen(QColor("#FFFFFF"), 2))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRect(c)
        p.setPen(QPen(QColor(255, 255, 255, 90), 1))
        for i in (1, 2):  # сетка «правило третей»
            p.drawLine(QPointF(c.x() + c.width() * i / 3, c.top()), QPointF(c.x() + c.width() * i / 3, c.bottom()))
            p.drawLine(QPointF(c.left(), c.y() + c.height() * i / 3), QPointF(c.right(), c.y() + c.height() * i / 3))
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor("#FFFFFF"))
        for pt in (c.topLeft(), c.topRight(), c.bottomLeft(), c.bottomRight()):
            p.drawEllipse(pt, 6, 6)

    def _corner(self, pos):
        c = self.crop_px()
        for name, pt in (("tl", c.topLeft()), ("tr", c.topRight()), ("bl", c.bottomLeft()), ("br", c.bottomRight())):
            if (pt - pos).manhattanLength() < 16:
                return name
        return "move" if c.contains(pos) else None

    def mousePressEvent(self, e):
        pos = e.position()
        part = self._corner(pos)
        if part:
            self._drag = (part, pos, self.crop_px())

    def mouseMoveEvent(self, e):
        pos = e.position()
        if not self._drag:
            part = self._corner(pos)
            self.setCursor(Qt.CursorShape.SizeAllCursor if part == "move" else
                           Qt.CursorShape.SizeFDiagCursor if part in ("tl", "br") else
                           Qt.CursorShape.SizeBDiagCursor if part in ("tr", "bl") else Qt.CursorShape.ArrowCursor)
            return
        part, start, c0 = self._drag
        ir = self.img_rect()
        dx, dy = pos.x() - start.x(), pos.y() - start.y()
        if part == "move":
            r = QRectF(c0).translated(dx, dy)
            r.moveLeft(max(ir.left(), min(r.left(), ir.right() - r.width())))
            r.moveTop(max(ir.top(), min(r.top(), ir.bottom() - r.height())))
        else:
            l, t, rr, b = c0.left(), c0.top(), c0.right(), c0.bottom()
            if "l" in part:
                l = min(rr - 20, max(ir.left(), l + dx))
            if "r" in part:
                rr = max(l + 20, min(ir.right(), rr + dx))
            if "t" in part:
                t = min(b - 20, max(ir.top(), t + dy))
            if "b" in part:
                b = max(t + 20, min(ir.bottom(), b + dy))
            r = QRectF(QPointF(l, t), QPointF(rr, b))
            if self.ratio:
                w = r.width()
                h = w / self.ratio
                if "t" in part:
                    r.setTop(r.bottom() - h)
                else:
                    r.setHeight(h)
                r = r.intersected(ir)
        self._set_px(r)
        self.update()

    def mouseReleaseEvent(self, e):
        self._drag = None


class CropDialog(QDialog):
    def __init__(self, parent, img, crop):
        super().__init__(parent)
        self.setWindowTitle("Обрезать фото")
        self.setObjectName("Dialog")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 16)
        t = QLabel("Обрезать фото")
        t.setObjectName("DialogTitle")
        lay.addWidget(t)
        self.view = CropView(img, crop)
        lay.addWidget(self.view, 1)
        row = QHBoxLayout()
        for label, ratio in (("Свободно", None), ("1:1", 1.0), ("4:3", 4 / 3), ("3:4", 3 / 4),
                             ("16:9", 16 / 9), ("9:16", 9 / 16)):
            b = QPushButton(label)
            b.setObjectName("Seg")
            b.setCheckable(True)
            b.setAutoExclusive(True)
            b.setChecked(ratio is None)
            b.clicked.connect(lambda _=False, r=ratio: self.view.set_ratio(r))
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        btns = QHBoxLayout()
        reset = QPushButton("Сбросить")
        reset.setObjectName("Ghost")
        reset.clicked.connect(lambda: (setattr(self.view, "crop", QRectF(0, 0, 1, 1)), self.view.update()))
        btns.addWidget(reset)
        btns.addStretch()
        cancel = QPushButton("Отмена")
        cancel.setObjectName("Ghost")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("Готово")
        ok.setObjectName("Pill")
        ok.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        lay.addLayout(btns)
        self.resize(720, 560)

    def result_crop(self):
        c = self.view.crop
        return [max(0.0, c.x()), max(0.0, c.y()), min(1.0, c.width()), min(1.0, c.height())]


# ───────────────────────────── Панель свойств ─────────────────────────────
HEAD_KINDS = [("Нет", "none"), ("Стрелка", "arrow"), ("Открытая", "open"), ("Круг", "circle"),
              ("Ромб", "diamond"), ("Черта", "bar")]


class Inspector(GlassBar):
    """Стеклянная панель справа: все настройки выбранной стрелки или стикера."""
    WIDTH = 336

    def __init__(self, win):
        super().__init__(win, radius=22)
        self.setObjectName("Inspector")
        self.item = None
        self.spin = {}
        from PySide6.QtWidgets import QScrollArea
        outer = QVBoxLayout(self)
        outer.setContentsMargins(self.PAD + 4, self.PAD + 2, self.PAD + 4, self.PAD + 6)
        outer.setSpacing(4)
        head = QHBoxLayout()
        self.title = QLabel("")
        self.title.setObjectName("DialogTitle")
        head.addWidget(self.title)
        head.addStretch()
        close = RoundButton("Скрыть панель", size=30)
        close.spec = ("svg", "close")
        win.icon_buttons.append(close)
        close.clicked.connect(self.hide_panel)
        head.addWidget(close)
        outer.addLayout(head)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll.setStyleSheet("QScrollArea{background:transparent;} QScrollArea > QWidget > QWidget{background:transparent;}")
        outer.addWidget(self.scroll, 1)
        self.hide()

    def auto(self):
        return self.win.settings.value("inspector_auto", "true") in (True, "true")

    def hide_panel(self):
        self.item = None
        self.hide()

    def place(self):
        area = self.win.editor_area
        top = self.win.glass.geometry().bottom() + 2
        content = self.scroll.widget().sizeHint().height() if self.scroll.widget() else 300
        h = min(area.height() - top - 10, content + 80)
        self.setGeometry(area.width() - self.WIDTH - 6, top, self.WIDTH, max(160, h))
        self.raise_()
        self.schedule()

    # ── конструктор полей ──
    def _section(self, lay, text):
        l = QLabel(text)
        l.setObjectName("Section")
        lay.addWidget(l)

    def _slider(self, lay, lo, hi, val, fn, suffix=""):
        from PySide6.QtWidgets import QSlider
        row = QHBoxLayout()
        sl = QSlider(Qt.Orientation.Horizontal)
        sl.setRange(lo, hi)
        sl.setValue(int(round(val)))
        lab = QLabel(f"{int(round(val))}{suffix}")
        lab.setObjectName("Muted")
        lab.setFixedWidth(42)
        lab.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        sl.valueChanged.connect(lambda v: (lab.setText(f"{v}{suffix}"), fn(v)))
        row.addWidget(sl, 1)
        row.addWidget(lab)
        lay.addLayout(row)
        return sl

    def _seg(self, lay, options, current, fn):
        grid = QGridLayout()
        grid.setSpacing(6)
        cols = 3 if len(options) != 4 else 2
        for i, (label, val) in enumerate(options):
            b = QPushButton(label)
            b.setObjectName("SegS")
            b.setCheckable(True)
            b.setChecked(val == current)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.clicked.connect(lambda _=False, v=val: (fn(v), self.rebuild()))
            grid.addWidget(b, i // cols, i % cols)
        lay.addLayout(grid)

    def _colors(self, lay, colors, current, fn, allow_none=False):
        cr = ColorRow(colors, current, allow_none=allow_none, per_row=7, size=24)
        cr.picked.connect(lambda c: (fn(c.name() if c is not None else None), self.rebuild()))
        lay.addWidget(cr)

    def _buttons(self, lay, items):
        grid = QGridLayout()
        grid.setSpacing(6)
        for i, (label, fn) in enumerate(items):
            b = QPushButton(label)
            b.setObjectName("SegS")
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.clicked.connect(fn)
            grid.addWidget(b, i // 2, i % 2)
        lay.addLayout(grid)

    def _combo(self, lay, label, options, current, fn):
        row = QHBoxLayout()
        l = QLabel(label)
        l.setObjectName("Muted")
        l.setFixedWidth(70)
        cb = QComboBox()
        for text, val in options:
            cb.addItem(text, val)
        idx = max(0, [v for _, v in options].index(current) if current in [v for _, v in options] else 0)
        cb.setCurrentIndex(idx)
        cb.setStyleSheet(f"background:{self.win.th['bg']};border:1px solid {self.win.th['border']};")
        cb.currentIndexChanged.connect(lambda i: fn(cb.itemData(i)))
        row.addWidget(l)
        row.addWidget(cb, 1)
        lay.addLayout(row)

    # ── содержимое ──
    def show_for(self, item):
        self.item = item
        self.rebuild()
        self.show()
        self.place()

    def rebuild(self):
        it = self.item
        if it is None:
            return
        pos = self.scroll.verticalScrollBar().value()
        box = QWidget()
        box.setMaximumWidth(self.WIDTH - 2 * self.PAD - 26)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(4, 0, 6, 8)
        lay.setSpacing(6)
        self.spin = {}
        if isinstance(it, ArrowSticker):
            self.title.setText("Стрелка")
            self.arrow_fields(lay, it)
        else:
            self.title.setText({"note": "Заметка", "photo": "Фото", "link": "Ссылка"}.get(it.d.get("type"), "Стикер"))
            self.sticker_fields(lay, it)
        lay.addStretch()
        self.scroll.setWidget(box)  # старое содержимое QScrollArea удаляет сам
        self.scroll.verticalScrollBar().setValue(pos)
        QTimer.singleShot(0, self.place)

    def arrow_fields(self, lay, a):
        d = a.d
        self._section(lay, "ЦВЕТ")
        self._colors(lay, ARROW_COLORS, d.get("color"), lambda c: a.set("color", c or "#3D5AFE"))
        self._section(lay, "ПРОЗРАЧНОСТЬ")
        self._slider(lay, 10, 100, float(d.get("opacity", 1)) * 100, lambda v: a.set("opacity", v / 100), "%")
        self._section(lay, "ТОЛЩИНА")
        self._slider(lay, 1, 14, float(d.get("width", 2.5)), lambda v: a.set("width", float(v)), " px")
        self._section(lay, "ЛИНИЯ")
        self._seg(lay, [("Сплошная", "solid"), ("Пунктир", "dash"), ("Точки", "dot")],
                  d.get("line", "solid"), lambda v: a.set("line", v))
        self._section(lay, "ФОРМА")
        self._seg(lay, [("Дуга", "curve"), ("Прямая", "straight"), ("Ступенька", "elbow")],
                  d.get("shape", "curve"), lambda v: a.set("shape", v))
        self._section(lay, "НАКОНЕЧНИКИ")
        self._combo(lay, "Начало", HEAD_KINDS, d.get("head_start", "none"), lambda v: a.set("head_start", v))
        self._combo(lay, "Конец", HEAD_KINDS, d.get("head_end", "arrow"), lambda v: a.set("head_end", v))
        lay.addSpacing(6)
        self._buttons(lay, [("Развернуть", a.reverse), ("Выпрямить", a.straighten),
                            ("Отцепить", a.detach), ("Удалить", a.delete)])

    def sticker_fields(self, lay, st):
        d = st.d
        typ = d.get("type")
        if typ != "photo":
            self._section(lay, "ЦВЕТ")
            self._colors(lay, STICKY_COLORS, d.get("color"), lambda c: st.set("color", c or STICKY_COLORS[0]))
        if typ == "note":
            self._section(lay, "ЦВЕТ ТЕКСТА")
            self._colors(lay, ["#1A1D24", "#FFFFFF", "#3C4043", "#E5484D", "#3D5AFE", "#30A46C"],
                         d.get("text_c"), lambda c: st.set("text_c", c), allow_none=True)
        self._section(lay, "ПРОЗРАЧНОСТЬ")
        self._slider(lay, 10, 100, float(d.get("opacity", 1)) * 100, lambda v: st.set("opacity", v / 100), "%")
        self._section(lay, "ФОРМА")
        self._seg(lay, [("Квадрат", "rect"), ("Скругл.", "round"), ("Капсула", "pill"), ("Овал", "ellipse")],
                  d.get("shape", "round"), lambda v: st.set("shape", v))
        if d.get("shape", "round") == "round":
            self._section(lay, "СКРУГЛЕНИЕ УГЛОВ")
            self._slider(lay, 0, 80, float(d.get("radius", 14)), lambda v: st.set("radius", v), " px")
        if typ == "photo":
            lay.addSpacing(2)
            self._buttons(lay, [("Обрезать…", st.crop), ("Без обрезки", lambda: (st.set("crop", None), st.crop_reset()))])
        self._section(lay, "РАМКА")
        self._colors(lay, ["#1A1D24", "#FFFFFF", "#3D5AFE", "#E5484D", "#F5B400", "#30A46C"],
                     d.get("border_c"), lambda c: st.set("border_c", c), allow_none=True)
        self._slider(lay, 0, 8, float(d.get("border_w", 1)), lambda v: st.set("border_w", v), " px")
        sh = QCheckBox("Тень")
        sh.setChecked(d.get("shadow", True))
        sh.toggled.connect(lambda on: st.set("shadow", on))
        lay.addWidget(sh)
        self._section(lay, "РАЗМЕР")
        row = QHBoxLayout()
        for key, label in (("w", "Ш"), ("h", "В")):
            l = QLabel(label)
            l.setObjectName("Muted")
            sb = QSpinBox()
            sb.setRange(60, 3000)
            sb.setValue(int(d[key]))
            sb.setSuffix(" px")
            sb.valueChanged.connect(lambda v, k=key: self._resize(st, k, v))
            self.spin[key] = sb
            row.addWidget(l)
            row.addWidget(sb, 1)
        lay.addLayout(row)
        self._section(lay, "ВЫРАВНИВАНИЕ НА СТРАНИЦЕ")
        self._buttons(lay, [("⇤ Слева", lambda: self._align(st, "left")), ("По центру", lambda: self._align(st, "center")),
                            ("Справа ⇥", lambda: self._align(st, "right")), ("Наверх экрана", lambda: self._align(st, "top")),
                            ("Вперёд", lambda: st.layer.bring(st, True)),
                            ("Назад", lambda: st.layer.bring(st, False))])
        lay.addSpacing(4)
        extra = [("Стрелка отсюда", st.start_arrow)]
        if typ == "note":
            extra.append(("Дедлайн…", st.edit_deadline))
        if typ == "link":
            extra.append(("Изменить ссылку…", st.edit_link))
        extra.append(("Удалить", st.delete))
        self._buttons(lay, extra)

    def _resize(self, st, key, v):
        st.d[key] = v
        if st.d.get("type") == "photo" and st.photo_ratio():
            m = Sticker.MARGIN
            if key == "w":
                st.d["h"] = round((v - 2 * m) / st.photo_ratio()) + 2 * m
            else:
                st.d["w"] = round((v - 2 * m) * st.photo_ratio()) + 2 * m
            other = "h" if key == "w" else "w"
            if other in self.spin:
                self.spin[other].blockSignals(True)
                self.spin[other].setValue(int(st.d[other]))
                self.spin[other].blockSignals(False)
        st.layer.place(st)
        st.layer.mark()

    def _align(self, st, where):
        vw = st.layer.viewport.width()
        margin = 56 - Sticker.MARGIN
        if where == "left":
            st.d["x"] = margin
        elif where == "center":
            st.d["x"] = (vw - st.d["w"]) / 2
        elif where == "right":
            st.d["x"] = vw - st.d["w"] - margin
        elif where == "top":
            st.d["y"] = st.layer.scroll() + 96
        st.layer.place(st)
        st.layer.mark()

    def refresh_values(self):
        it = self.item
        if it is None or not self.isVisible() or isinstance(it, ArrowSticker):
            return
        for key, sb in self.spin.items():
            sb.blockSignals(True)
            sb.setValue(int(it.d[key]))
            sb.blockSignals(False)


# ───────────────────────────── Обновления через GitHub ─────────────────────────────
def version_tuple(v):
    nums = [int(x) for x in re.findall(r"\d+", v or "")[:3]]
    return tuple(nums + [0] * (3 - len(nums)))


class Updater(QObject):
    """Проверяет последнюю версию на GitHub и обновляет программу одним кликом."""
    def __init__(self, win):
        super().__init__(win)
        from PySide6.QtNetwork import QNetworkAccessManager
        self.win = win
        self.nam = QNetworkAccessManager(self)
        self.latest = None  # (tag, url_zip, url_page)

    @staticmethod
    def can_self_update():
        return sys.platform == "win32" and getattr(sys, "frozen", False)

    def _request(self, url):
        from PySide6.QtNetwork import QNetworkRequest
        req = QNetworkRequest(QUrl(url))
        req.setRawHeader(b"User-Agent", b"Folio-Updater")
        req.setRawHeader(b"Accept", b"application/vnd.github+json")
        return req

    def check(self, manual=False):
        reply = self.nam.get(self._request(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"))
        reply.finished.connect(lambda: self._checked(reply, manual))

    def _checked(self, reply, manual):
        from PySide6.QtNetwork import QNetworkReply
        try:
            if reply.error() != QNetworkReply.NetworkError.NoError:
                if manual:
                    QMessageBox.information(self.win, APP_NAME, "Не получилось проверить обновления.\n"
                                                                "Проверь интернет и попробуй ещё раз.")
                return
            data = json.loads(bytes(reply.readAll()).decode("utf-8"))
            tag = data.get("tag_name", "")
            zip_url = next((a.get("browser_download_url") for a in data.get("assets", [])
                            if a.get("name", "").lower().endswith("windows.zip")), None)
            if version_tuple(tag) > version_tuple(APP_VERSION) and zip_url:
                self.latest = (tag, zip_url, data.get("html_url", ""))
                self.win.show_update_available(tag)
                if manual:
                    self.ask_install()
            elif manual:
                QMessageBox.information(self.win, APP_NAME, f"У тебя последняя версия — {APP_VERSION} ✓")
        except Exception as e:
            if manual:
                QMessageBox.information(self.win, APP_NAME, f"Не получилось проверить обновления:\n{e}")
        finally:
            reply.deleteLater()

    def ask_install(self):
        if not self.latest:
            return
        tag, zip_url, page = self.latest
        if not self.can_self_update():
            QDesktopServices.openUrl(QUrl(page or f"https://github.com/{GITHUB_REPO}/releases/latest"))
            return
        box = QMessageBox(self.win)
        box.setWindowTitle(APP_NAME)
        box.setIcon(QMessageBox.Icon.NoIcon)
        box.setText(f"Доступна новая версия {tag.lstrip('v')}")
        box.setInformativeText("Folio скачает её, закроется на пару секунд и откроется уже обновлённым.\n"
                               "Твои документы, проекты и настройки останутся на месте.")
        go = box.addButton("Обновить", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("Позже", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is go:
            self.download()

    def download(self):
        tag, zip_url, _ = self.latest
        reply = self.nam.get(self._request(zip_url))
        reply.downloadProgress.connect(lambda got, total: self.win.show_any_toast(
            f"Скачиваю обновление… {int(got * 100 / total)}%" if total > 0 else "Скачиваю обновление…"))
        reply.finished.connect(lambda: self._downloaded(reply))

    def _downloaded(self, reply):
        import tempfile
        from PySide6.QtNetwork import QNetworkReply
        try:
            if reply.error() != QNetworkReply.NetworkError.NoError:
                QMessageBox.warning(self.win, APP_NAME, "Обновление не скачалось. Попробуй позже.")
                return
            tmp = Path(tempfile.mkdtemp(prefix="folio-update-"))
            zpath = tmp / "update.zip"
            zpath.write_bytes(bytes(reply.readAll()))
            with zipfile.ZipFile(zpath) as z:
                z.extractall(tmp / "x")
            exe = next((p for p in (tmp / "x").rglob("Folio.exe")), None)
            if exe is None:
                QMessageBox.warning(self.win, APP_NAME, "В обновлении не нашлось Folio.exe.")
                return
            src = exe.parent
            dst = Path(sys.executable).parent
            script = tmp / "update.ps1"
            script.write_text(
                "param([int]$Pid_, [string]$Src, [string]$Dst)\r\n"
                "try { Wait-Process -Id $Pid_ -Timeout 30 -ErrorAction SilentlyContinue } catch {}\r\n"
                "Start-Sleep -Milliseconds 500\r\n"
                "robocopy $Src $Dst /MIR /R:3 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null\r\n"
                "Start-Process (Join-Path $Dst 'Folio.exe')\r\n", encoding="utf-8-sig")
            if not self.win.prepare_quit():
                return
            subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                              "-File", str(script), "-Pid_", str(os.getpid()), "-Src", str(src), "-Dst", str(dst)],
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            QApplication.quit()
        except Exception as e:
            QMessageBox.warning(self.win, APP_NAME, f"Не получилось установить обновление:\n{e}")
        finally:
            reply.deleteLater()


# ───────────────────────────── Главное окно ─────────────────────────────
class Folio(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings(APP_NAME, APP_NAME)
        self.path = None
        self.fmt = "docx"
        self.encoding = "utf-8"
        self.newline = "\r\n"
        self.icon_buttons = []
        self._updating = False

        self.theme_mode = self.settings.value("theme", "system")
        if self.theme_mode not in ("light", "dark", "system"):
            self.theme_mode = "system"
        self.theme = self.resolve_theme()
        try:
            QGuiApplication.styleHints().colorSchemeChanged.connect(self._system_scheme_changed)
        except Exception:
            pass
        global DEFAULT_FONT, DEFAULT_SIZE
        DEFAULT_FONT = self.settings.value("doc_font", DEFAULT_FONT) or DEFAULT_FONT
        try:
            DEFAULT_SIZE = float(self.settings.value("doc_size", DEFAULT_SIZE))
        except (TypeError, ValueError):
            pass

        self.setWindowIcon(app_icon())
        self.resize(1240, 860)
        self.setMinimumSize(1060, 560)
        geo = self.settings.value("geometry")
        if geo:
            self.restoreGeometry(geo)

        self.ui_font = "Segoe UI"
        self.th = THEMES[self.theme]
        self.stack = QStackedWidget()
        self.stack.setObjectName("Root")
        self.setCentralWidget(self.stack)

        self.browser = BrowserPage(self)
        self.stack.addWidget(self.browser)

        self.editor_page = QWidget()
        self.editor_page.setObjectName("Root")
        v = QVBoxLayout(self.editor_page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        v.addWidget(self._build_appbar())
        self._build_fmtbar()
        v.addWidget(self._build_findbar())
        v.addWidget(self._build_editor_area(), 1)
        v.addWidget(self._build_statusbar())
        self.stack.addWidget(self.editor_page)

        self._build_shortcuts()
        self.apply_theme()
        self.new_document(ask=False)
        self.updater = Updater(self)
        if self.settings.value("auto_update", "true") in (True, "true"):
            QTimer.singleShot(5000, lambda: self.updater.check(False))
        last = self.settings.value("last_folder", "")
        self.browser.navigate(last if last else self.browser.root)
        self.stack.setCurrentWidget(self.browser)

    # ── UI ──
    def _btn(self, tip, spec, slot=None, checkable=False):
        b = RoundButton(tip, checkable)
        b.spec = spec
        if slot:
            (b.toggled if checkable else b.clicked).connect(slot)
        self.icon_buttons.append(b)
        return b

    def _build_appbar(self):
        bar = QWidget()
        bar.setObjectName("AppBar")
        bar.setFixedHeight(64)
        h = QHBoxLayout(bar)
        h.setContentsMargins(16, 10, 16, 6)
        h.setSpacing(4)

        h.addWidget(self._btn("К проектам  (Alt+←)", ("svg", "back"), self.go_home))
        h.addSpacing(4)

        tcol = QVBoxLayout()
        tcol.setSpacing(0)
        self.title_label = QLabel("Без названия")
        self.title_label.setObjectName("DocTitle")
        self.sub_label = QLabel("Новый документ")
        self.sub_label.setObjectName("DocSub")
        tcol.addWidget(self.title_label)
        tcol.addWidget(self.sub_label)
        h.addLayout(tcol)
        h.addSpacing(12)

        h.addWidget(self._btn("Новый документ  (Ctrl+N)", ("svg", "new"), lambda: self.new_document()))
        h.addWidget(self._btn("Открыть  (Ctrl+O)", ("svg", "open"), self.open_dialog))
        h.addStretch()

        self.undo_btn = self._btn("Отменить  (Ctrl+Z)", ("svg", "undo"), lambda: self.editor.undo())
        self.redo_btn = self._btn("Повторить  (Ctrl+Y)", ("svg", "redo"), lambda: self.editor.redo())
        h.addWidget(self.undo_btn)
        h.addWidget(self.redo_btn)
        h.addWidget(self._btn("Найти и заменить  (Ctrl+F)", ("svg", "find"), self.toggle_find))
        self.theme_btn = self._btn("Тёмная / светлая тема", ("svg", "moon"), self.toggle_theme)
        h.addWidget(self.theme_btn)
        h.addWidget(self._btn("Внешний вид: цвета, страница, шрифт", ("svg", "tune"), self.open_settings))

        self.more_btn = self._btn("Ещё", ("svg", "more"))
        self.more_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.more_btn.setObjectName("NoArrow")
        menu = styled_menu(self)
        menu.addAction("Сохранить как…", self.save_as).setShortcut("Ctrl+Shift+S")
        self.export_menu_ref = export_menu(menu, self.export_as, self.th)
        menu.addMenu(self.export_menu_ref)
        menu.addAction("Печать…", self.print_doc).setShortcut("Ctrl+P")
        menu.addSeparator()
        menu.addAction("Очистить форматирование", self.clear_formatting)
        menu.addAction("Вставить картинку…", self.insert_image)
        menu.addAction("Вставить таблицу…", self.insert_table)
        self.more_btn.setMenu(menu)
        h.addWidget(self.more_btn)
        h.addSpacing(8)

        self.save_btn = QPushButton("Сохранить")
        self.save_btn.setObjectName("Pill")
        self.save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_btn.setToolTip("Ctrl+S")
        self.save_btn.clicked.connect(self.save)
        h.addWidget(self.save_btn)
        return bar

    def _sep(self):
        s = QFrame()
        s.setObjectName("Sep")
        s.setFixedSize(1, 22)
        return s

    def _build_fmtbar(self):
        bar = GlassBar(self)
        self.glass = bar
        pad = GlassBar.PAD
        h = QHBoxLayout(bar)
        h.setContentsMargins(pad + 10, pad + 1, pad + 8, pad + 5)
        h.setSpacing(2)

        self.style_box = QComboBox()
        self.style_box.addItems(["Обычный текст", "Заголовок 1", "Заголовок 2", "Заголовок 3"])
        self.style_box.setFixedWidth(132)
        self.style_box.setToolTip("Стиль абзаца")
        self.style_box.activated.connect(self.set_block_style)
        h.addWidget(self.style_box)

        self.font_box = QFontComboBox()
        self.font_box.setFixedWidth(150)
        self.font_box.setMaxVisibleItems(16)
        self.font_box.setToolTip("Шрифт")
        self.font_box.textActivated.connect(self.set_font_family)
        h.addWidget(self.font_box)

        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        self.size_box.addItems([str(s) for s in (8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72)])
        self.size_box.setFixedWidth(68)
        self.size_box.setToolTip("Размер")
        self.size_box.textActivated.connect(self.set_font_size)
        h.addWidget(self.size_box)
        h.addSpacing(4)
        h.addWidget(self._sep())
        h.addSpacing(4)

        self.bold_btn = self._btn("Жирный  (Ctrl+B)", ("glyph", "B", {"bold": True}), self.set_bold, True)
        self.italic_btn = self._btn("Курсив  (Ctrl+I)", ("glyph", "I", {"italic": True}), self.set_italic, True)
        self.under_btn = self._btn("Подчёркнутый  (Ctrl+U) · стрелка — волна, точки, цвет линии",
                                   ("glyph", "U", {"underline": True}), self.set_underline, True)
        self.under_btn.setObjectName("Split")
        self.under_btn.setFixedSize(48, 36)
        self.under_btn.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        self.under_btn.setMenu(self._underline_menu())
        self.strike_btn = self._btn("Зачёркнутый", ("glyph", "S", {"strike": True}), self.set_strike, True)
        for b in (self.bold_btn, self.italic_btn, self.under_btn, self.strike_btn):
            h.addWidget(b)

        self.color_btn = self._btn("Цвет текста", ("glyph", "A", {"bar": "#E5484D"}))
        self.color_btn.setObjectName("NoArrow")
        self.color_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        cm = SwatchMenu(self)
        cm.picked.connect(self.set_text_color)
        self.color_btn.setMenu(cm)
        h.addWidget(self.color_btn)

        self.hl_btn = self._btn("Выделение цветом", ("svg", "marker"))
        self.hl_btn.setObjectName("NoArrow")
        self.hl_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        hm = SwatchMenu(self, highlight=True)
        hm.picked.connect(self.set_highlight)
        self.hl_btn.setMenu(hm)
        h.addWidget(self.hl_btn)
        h.addWidget(self._btn("Ссылка  (Ctrl+K)", ("svg", "link"), self.insert_link))

        h.addSpacing(4)
        h.addWidget(self._sep())
        h.addSpacing(4)
        self.bullet_btn = self._btn("Маркированный список", ("svg", "bullets"), lambda c: self.toggle_list("bullet", c), True)
        self.number_btn = self._btn("Нумерованный список", ("svg", "numbers"), lambda c: self.toggle_list("number", c), True)
        self.check_btn = self._btn("Список дел — с галочками", ("svg", "checklist"), lambda c: self.toggle_list("check", c), True)
        h.addWidget(self.bullet_btn)
        h.addWidget(self.number_btn)
        h.addWidget(self.check_btn)
        h.addSpacing(4)
        h.addWidget(self._sep())
        h.addSpacing(4)

        self.align_btns = {}
        for key, tip, flag in (("left", "По левому краю", Qt.AlignmentFlag.AlignLeft),
                               ("center", "По центру", Qt.AlignmentFlag.AlignHCenter),
                               ("right", "По правому краю", Qt.AlignmentFlag.AlignRight),
                               ("justify", "По ширине", Qt.AlignmentFlag.AlignJustify)):
            b = self._btn(tip, ("svg", key), checkable=True)
            b.clicked.connect(lambda _=False, f=flag: self.set_align(f))
            self.align_btns[flag] = b
            h.addWidget(b)

        h.addSpacing(4)
        h.addWidget(self._sep())
        h.addSpacing(6)
        self.add_btn = QToolButton()
        self.add_btn.setObjectName("AddBtn")
        self.add_btn.setToolTip("Добавить: заметку, список дел, дедлайн, фото, ссылку…")
        self.add_btn.setFixedSize(36, 36)
        self.add_btn.setIconSize(QSize(20, 20))
        self.add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.add_btn.clicked.connect(self.show_add_menu)
        h.addWidget(self.add_btn)
        return bar

    def _underline_menu(self):
        m = styled_menu(self)
        US = QTextCharFormat.UnderlineStyle
        for label, style in (("Обычное", US.SingleUnderline), ("Волнистое", US.WaveUnderline),
                             ("Пунктир", US.DashUnderline), ("Точками", US.DotLine),
                             ("Штрих-пунктир", US.DashDotLine)):
            m.addAction(label, lambda st=style: self.set_underline_style(st))
        m.addSeparator()
        cm = SwatchMenu(m, reset=True)
        cm.setTitle("Цвет линии")
        cm.picked.connect(self.set_underline_color)
        m.addMenu(cm)
        return m

    def show_add_menu(self):
        th = self.th
        m = styled_menu(self)
        m.addSection("Стикеры — видны только в Folio")
        m.addAction(make_icon(("svg", "sticky"), th), "Заметка", lambda: self.stickers.add_note("note"))
        m.addAction(make_icon(("svg", "checklist"), th), "Заметка — список дел", lambda: self.stickers.add_note("todo"))
        m.addAction(make_icon(("svg", "clock"), th), "Дедлайн", lambda: self.stickers.add_note("deadline"))
        m.addAction(make_icon(("svg", "image"), th), "Фото-стикер…", self.add_photo_sticker)
        m.addAction(make_icon(("svg", "link"), th), "Ссылка-стикер…", self.stickers.add_link)
        m.addAction(make_icon(("svg", "arrow"), th), "Стрелка", lambda: self.stickers.add_arrow())
        m.addSection("В текст документа")
        m.addAction(make_icon(("svg", "checklist"), th), "Список дел", lambda: self.toggle_list("check", True))
        m.addAction(make_icon(("svg", "image"), th), "Картинка…", self.insert_image)
        m.addAction(make_icon(("svg", "table"), th), "Таблица…", self.insert_table)
        m.addAction(make_icon(("svg", "link"), th), "Ссылка…  (Ctrl+K)", self.insert_link)
        m.addSeparator()
        m.addAction(make_icon(("svg", "fontplus"), th), "Добавить шрифт из файла…", self.add_font_file)
        m.exec(self.add_btn.mapToGlobal(self.add_btn.rect().bottomLeft() + QPoint(0, 6)))

    def add_photo_sticker(self):
        start = self.settings.value("last_img_dir", str(Path.home() / "Pictures"))
        paths, _ = QFileDialog.getOpenFileNames(self, "Фото-стикер", start,
                                                "Изображения (*.png *.jpg *.jpeg *.gif *.bmp *.webp)")
        for i, p in enumerate(paths):
            self.settings.setValue("last_img_dir", str(Path(p).parent))
            self.stickers.add_photo(p)

    def add_font_file(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Добавить шрифт", str(Path.home() / "Downloads"),
                                                "Шрифты (*.ttf *.otf *.ttc)")
        added = []
        for p in paths:
            dst = fonts_dir() / Path(p).name
            try:
                if Path(p).resolve() != dst.resolve():
                    shutil.copy2(p, dst)
            except OSError as e:
                QMessageBox.warning(self, APP_NAME, f"Не получилось скопировать шрифт:\n{e}")
                continue
            fid = QFontDatabase.addApplicationFont(str(dst))
            if fid < 0:
                QMessageBox.warning(self, APP_NAME, f"Файл «{Path(p).name}» не похож на шрифт.")
                continue
            added += QFontDatabase.applicationFontFamilies(fid)
        if added:
            self.font_box.setFontFilters(self.font_box.fontFilters())  # обновить список
            self.font_box.setCurrentFont(QFont(added[0]))
            self.show_toast(f"Шрифт «{added[0]}» добавлен ✓")

    def _build_findbar(self):
        wrap = QWidget()
        wrap.setObjectName("FmtWrap")
        outer = QHBoxLayout(wrap)
        outer.setContentsMargins(16, 0, 16, 10)
        outer.addStretch()
        bar = QFrame()
        bar.setObjectName("FmtBar")
        h = QHBoxLayout(bar)
        h.setContentsMargins(12, 5, 6, 5)
        h.setSpacing(4)
        self.find_edit = QLineEdit()
        self.find_edit.setPlaceholderText("Найти")
        self.find_edit.setFixedWidth(200)
        self.find_edit.returnPressed.connect(lambda: self.find_next())
        self.find_edit.textChanged.connect(lambda: self.find_edit.setProperty("error", False) or self._repolish(self.find_edit))
        self.repl_edit = QLineEdit()
        self.repl_edit.setPlaceholderText("Заменить на")
        self.repl_edit.setFixedWidth(200)
        self.repl_edit.returnPressed.connect(self.replace_one)
        h.addWidget(self.find_edit)
        h.addWidget(self._btn("Предыдущее  (Shift+Enter)", ("svg", "up"), lambda: self.find_next(True)))
        h.addWidget(self._btn("Следующее  (Enter)", ("svg", "down"), lambda: self.find_next()))
        h.addSpacing(6)
        h.addWidget(self.repl_edit)
        r1 = QPushButton("Заменить")
        r1.setObjectName("Ghost")
        r1.clicked.connect(self.replace_one)
        r2 = QPushButton("Заменить все")
        r2.setObjectName("Ghost")
        r2.clicked.connect(self.replace_all)
        h.addWidget(r1)
        h.addWidget(r2)
        h.addWidget(self._btn("Закрыть  (Esc)", ("svg", "close"), self.toggle_find))
        outer.addWidget(bar)
        outer.addStretch()
        self.findbar = wrap
        wrap.hide()
        sc = QShortcut(QKeySequence("Shift+Return"), self.find_edit)
        sc.activated.connect(lambda: self.find_next(True))
        return wrap

    def _build_editor_area(self):
        area = QWidget()
        area.setObjectName("EditorArea")
        self.editor_area = area
        area.installEventFilter(self)
        h = QHBoxLayout(area)
        h.setContentsMargins(16, 0, 16, 0)
        h.addStretch()
        self.editor = Editor()
        self.editor.setObjectName("Editor")
        self.editor.setMaximumWidth(880)
        self.editor.setMinimumWidth(560)
        self.editor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.editor.setPlaceholderText("Начни писать…")
        self.editor.setAcceptRichText(True)
        self.editor.setTabStopDistance(36)
        self.editor.viewport().setAutoFillBackground(False)
        self.editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.editor.customContextMenuRequested.connect(self.context_menu)
        self.editor.fileDropped.connect(self.open_path_checked)
        self.editor.currentCharFormatChanged.connect(lambda _: self.sync_toolbar())
        self.editor.cursorPositionChanged.connect(self.sync_toolbar)
        self.editor.undoAvailable.connect(lambda a: self.undo_btn.setEnabled(a))
        self.editor.redoAvailable.connect(lambda a: self.redo_btn.setEnabled(a))
        self.editor.textChanged.connect(lambda: self.count_timer.start())
        h.addWidget(self.editor, 1)
        h.addStretch()
        self.glass.setParent(area)
        self.glass.raise_()
        self.editor.verticalScrollBar().valueChanged.connect(lambda _: (self.glass.schedule(), self.inspector.schedule()))
        self.editor.textChanged.connect(self.glass.schedule)
        self.editor.selectionChanged.connect(self.glass.schedule)
        self.stickers = StickerLayer(self, self.editor)
        self.inspector = Inspector(self)
        self.inspector.setParent(area)
        self.count_timer = QTimer(self)
        self.count_timer.setSingleShot(True)
        self.count_timer.setInterval(250)
        self.count_timer.timeout.connect(self.update_counts)
        return area

    def _build_statusbar(self):
        bar = QWidget()
        bar.setObjectName("StatusBar")
        bar.setFixedHeight(40)
        h = QHBoxLayout(bar)
        h.setContentsMargins(24, 6, 24, 8)
        self.count_label = QLabel("0 слов")
        self.count_label.setObjectName("Muted")
        self.toast = QLabel("")
        self.toast.setObjectName("Toast")
        self.count_label.setMinimumWidth(260)
        self.toast.setMinimumWidth(260)
        self.toast.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.format_chip = QLabel("DOCX")
        self.format_chip.setObjectName("Chip")
        h.addWidget(self.count_label)
        h.addStretch()
        h.addWidget(self.toast)
        h.addStretch()
        right = QWidget()
        right.setMinimumWidth(260)
        rl = QHBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.addStretch()
        rl.addWidget(self.format_chip)
        h.addWidget(right)
        self.toast_timer = QTimer(self)
        self.toast_timer.setSingleShot(True)
        self.toast_timer.timeout.connect(lambda: self.toast.setText(""))
        return bar

    def _build_shortcuts(self):
        def glob(keys, fn):
            s = QShortcut(QKeySequence(keys), self)
            s.activated.connect(fn)

        def sc(keys, fn):  # работает только на экране редактора
            s = QShortcut(QKeySequence(keys), self.editor_page)
            s.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
            s.activated.connect(fn)
        glob("Ctrl+N", self.new_shortcut)
        glob("Ctrl+O", self.open_dialog)
        sc("Alt+Left", self.go_home)
        sc("Ctrl+K", self.insert_link)
        sc("Ctrl+S", self.save)
        sc("Ctrl+Shift+S", self.save_as)
        sc("Ctrl+P", self.print_doc)
        sc("Ctrl+B", self.bold_btn.click)
        sc("Ctrl+I", self.italic_btn.click)
        sc("Ctrl+U", self.under_btn.click)
        sc("Ctrl+F", lambda: self.toggle_find(True))
        sc("Ctrl+H", lambda: self.toggle_find(True))
        sc("Ctrl+Y", lambda: self.editor.redo())
        sc("Ctrl+Alt+1", lambda: self.set_block_style(1))
        sc("Ctrl+Alt+2", lambda: self.set_block_style(2))
        sc("Ctrl+Alt+3", lambda: self.set_block_style(3))
        sc("Ctrl+Alt+0", lambda: self.set_block_style(0))
        sc("Ctrl+E", lambda: self.set_align(Qt.AlignmentFlag.AlignHCenter))
        sc("Ctrl+L", lambda: self.set_align(Qt.AlignmentFlag.AlignLeft))
        sc("Ctrl+R", lambda: self.set_align(Qt.AlignmentFlag.AlignRight))
        sc("Ctrl+J", lambda: self.set_align(Qt.AlignmentFlag.AlignJustify))
        sc("Ctrl+Shift+8", lambda: self.bullet_btn.click())
        sc("Ctrl+Shift+7", lambda: self.number_btn.click())
        bs = QShortcut(QKeySequence("Alt+Left"), self.browser)
        bs.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        bs.activated.connect(self.browser.go_up)
        esc = QShortcut(QKeySequence("Escape"), self.editor_page)
        esc.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        esc.activated.connect(lambda: self.findbar.isVisible() and self.toggle_find(False))

    # ── Тема ──
    def eventFilter(self, obj, ev):
        if obj is getattr(self, "editor_area", None) and ev.type() in (QEvent.Type.Resize, QEvent.Type.Show):
            self.place_glass()
        return super().eventFilter(obj, ev)

    def place_glass(self):
        if self.inspector.isVisible():
            QTimer.singleShot(0, self.inspector.place)
        g, area = self.glass, self.editor_area
        hint = g.sizeHint()
        w = min(hint.width(), area.width())
        g.setGeometry(max(0, (area.width() - w) // 2), 0, w, hint.height())
        g.raise_()
        g.schedule()

    def resolve_theme(self):
        if self.theme_mode in ("light", "dark"):
            return self.theme_mode
        try:
            return "dark" if QGuiApplication.styleHints().colorScheme() == Qt.ColorScheme.Dark else "light"
        except Exception:
            return "light"

    def _system_scheme_changed(self, *_):
        if self.theme_mode == "system":
            self.theme = self.resolve_theme()
            self.apply_theme()

    def current_theme(self):
        s = self.settings
        dark = self.theme == "dark"
        return build_theme(dark, s.value("accent", "") or None, s.value("tint", "false") in (True, "true"),
                           s.value(f"page_{self.theme}", "") or None, s.value(f"text_{self.theme}", "") or None)

    def apply_theme(self):
        th = self.current_theme()
        self.th = th
        for b in self.icon_buttons:
            if b.spec:
                b.setIcon(make_icon(b.spec, th))
        for tb in (self.theme_btn, self.browser.theme_btn):
            tb.setIcon(make_icon(("svg", "sun" if self.theme == "dark" else "moon"), th))
        self.browser.create_btn.setIcon(QIcon(svg_pixmap(SVG["plus"], th["accent_text"], size=18, stroke=2.2)))
        self._update_color_icon()
        self.browser.grid.viewport().update()
        self.add_btn.setIcon(QIcon(svg_pixmap(SVG["plus"], th["accent_text"], size=20, stroke=2.4)))
        if hasattr(self, "stickers"):
            self.stickers.refresh_styles()
        if hasattr(self, "glass"):
            self.glass.schedule()
            QTimer.singleShot(0, self.place_glass)
        if hasattr(self, "export_menu_ref"):
            self.export_menu_ref.setIcon(make_icon(("svg", "export"), th))

        # переключатель (вкл/выкл) с кружком
        tmp = Path(os.environ.get("TEMP", "/tmp"))
        sw = {}
        for state, track, knob_x in (("off", th["border"], 3), ("on", th["accent"], 19)):
            pm = QPixmap(76, 44)
            pm.fill(Qt.GlobalColor.transparent)
            p = QPainter(pm)
            p.setRenderHint(QPainter.RenderHint.Antialiasing)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(track))
            p.drawRoundedRect(QRectF(0, 0, 76, 44), 22, 22)
            p.setBrush(QColor("#FFFFFF"))
            p.drawEllipse(QRectF(knob_x * 2, 6, 32, 32))
            p.end()
            f = tmp / f"folio_switch_{state}_{QColor(track).name()[1:]}.png"
            pm.save(str(f))
            sw[state] = str(f).replace("\\", "/")
        self._switch = sw

        # стрелка для выпадающих списков
        arrow = Path(os.environ.get("TEMP", "/tmp")) / f"folio_arrow_{self.theme}.png"
        svg_pixmap(SVG["chev"], th["muted"], size=14, dpr=2).save(str(arrow))
        arrow_url = str(arrow).replace("\\", "/")

        st = self.settings
        self.rf = float(st.value("ui_round", 1.0))
        ui_size = float(st.value("ui_font", 10))
        self.editor.setMaximumWidth(int(st.value("page_width", 880)))
        sheet = f"""
        * {{ font-family: "Segoe UI Variable Text", "Segoe UI", "Inter", sans-serif; font-size: {ui_size:g}pt; color: {th['text']}; }}
        #Root, #AppBar, #FmtWrap, #EditorArea, #StatusBar {{ background: {th['bg']}; }}
        #DocTitle {{ font-size: 12pt; font-weight: 600; }}
        #DocSub, #Muted {{ color: {th['muted']}; font-size: 9pt; }}
        #Toast {{ color: {th['accent']}; font-size: 9pt; font-weight: 600; }}
        #Chip {{ background: {th['surface']}; border: 1px solid {th['border']}; border-radius: 11px;
                 padding: 2px 10px; color: {th['muted']}; font-size: 8.5pt; font-weight: 600; }}
        #FmtBar {{ background: {th['surface']}; border: 1px solid {th['border']}; border-radius: 23px; }}
        #Sep {{ background: {th['border']}; }}
        QToolButton {{ background: transparent; border: none; border-radius: 18px; }}
        QToolButton:hover {{ background: {th['hover']}; }}
        QToolButton:pressed {{ background: {th['soft']}; }}
        QToolButton:checked {{ background: {th['soft']}; }}
        QToolButton#NoArrow::menu-indicator {{ image: none; width: 0px; }}
        QPushButton#Pill {{ background: {th['accent']}; color: {th['accent_text']}; border: none; border-radius: 18px;
                            padding: 0 20px; min-height: 36px; font-weight: 600; }}
        QPushButton#Pill:hover {{ background: {QColor(th['accent']).lighter(112).name()}; }}
        QPushButton#Ghost {{ background: transparent; border: 1px solid {th['border']}; border-radius: 15px;
                             padding: 0 14px; min-height: 30px; }}
        QPushButton#Ghost:hover {{ background: {th['hover']}; }}
        QComboBox, QFontComboBox, QSpinBox, QLineEdit {{
            background: transparent; border: 1px solid transparent; border-radius: 12px;
            padding: 4px 8px; min-height: 26px; selection-background-color: {th['sel']}; }}
        QComboBox:hover, QFontComboBox:hover {{ background: {th['hover']}; }}
        QLineEdit, QSpinBox {{ background: {th['bg']}; border: 1px solid {th['border']}; }}
        QLineEdit:focus, QSpinBox:focus {{ border: 1px solid {th['accent']}; }}
        QLineEdit[error="true"] {{ border: 1px solid {th['danger']}; }}
        QComboBox::drop-down {{ border: none; width: 18px; }}
        QComboBox::down-arrow {{ image: url("{arrow_url}"); width: 14px; height: 14px; }}
        QComboBox QAbstractItemView {{ background: {th['surface']}; border: 1px solid {th['border']};
            border-radius: 10px; padding: 4px; outline: none; selection-background-color: {th['soft']};
            selection-color: {th['text']}; }}
        QSpinBox::up-button, QSpinBox::down-button {{ width: 0px; border: none; }}
        #Editor {{ font-family: "{DEFAULT_FONT}"; font-size: {DEFAULT_SIZE}pt; background: {th['page']}; border: 1px solid {th['border']}; border-top-left-radius: 18px;
                   border-top-right-radius: 18px; border-bottom: none; color: {th['doc_text']};
                   selection-background-color: {th['sel']}; selection-color: {th['doc_text']}; }}
        #GlassBar, #Inspector {{ background: transparent; border: none; }}
        QSlider::groove:horizontal {{ height: 4px; background: {th['border']}; border-radius: 2px; }}
        QSlider::sub-page:horizontal {{ background: {th['accent']}; border-radius: 2px; }}
        QSlider::handle:horizontal {{ background: #FFFFFF; border: 2px solid {th['accent']}; width: 14px; height: 14px;
                                      margin: -7px 0; border-radius: 9px; }}
        QToolButton#AddBtn {{ background: {th['accent']}; border-radius: 18px; }}
        QToolButton#AddBtn:hover {{ background: {QColor(th['accent']).lighter(112).name()}; }}
        QToolButton#Split {{ padding-right: 12px; border-radius: 18px; }}
        QToolButton#Split::menu-button {{ border: none; width: 14px; border-top-right-radius: 18px;
                                          border-bottom-right-radius: 18px; }}
        QToolButton#Split::menu-arrow {{ image: url("{arrow_url}"); width: 10px; height: 10px; }}
        QPushButton#DeadlineChip {{ text-align: left; }}
        #Section {{ color: {th['muted']}; font-size: 8.5pt; font-weight: 700; letter-spacing: 0.6px; padding-top: 8px; }}
        QPushButton#Seg {{ background: transparent; border: 1px solid {th['border']}; border-radius: 17px;
                           padding: 0 16px; min-height: 34px; }}
        QPushButton#Seg:hover {{ background: {th['hover']}; }}
        QPushButton#SegS {{ background: transparent; border: 1px solid {th['border']}; border-radius: 15px;
                            padding: 0 8px; min-height: 30px; }}
        QPushButton#SegS:hover {{ background: {th['hover']}; }}
        QPushButton#SegS:checked {{ background: {th['soft']}; border: 1px solid {th['accent']}; color: {th['accent']}; font-weight: 600; }}
        QPushButton#Seg:checked {{ background: {th['soft']}; border: 1px solid {th['accent']}; color: {th['accent']}; font-weight: 600; }}
        QCheckBox {{ spacing: 10px; }}
        QCheckBox::indicator {{ width: 38px; height: 22px; image: url("{sw['off']}"); }}
        QCheckBox::indicator:checked {{ image: url("{sw['on']}"); }}
        QMenu {{ background: {th['surface']}; border: 1px solid {th['border']}; border-radius: 14px; padding: 6px; }}
        QMenu::item {{ padding: 8px 28px 8px 14px; border-radius: 9px; }}
        QMenu::item:selected {{ background: {th['hover']}; }}
        QMenu::separator {{ height: 1px; background: {th['border']}; margin: 5px 8px; }}
        #SwatchBox {{ background: transparent; }}
        QPushButton#SwatchReset {{ background: transparent; border: 1px solid {th['border']}; border-radius: 13px;
                                   padding: 4px 10px; color: {th['muted']}; }}
        QPushButton#SwatchReset:hover {{ background: {th['hover']}; }}
        QScrollBar:vertical {{ background: transparent; width: 12px; margin: 18px 2px 4px 0; }}
        QScrollBar::handle:vertical {{ background: {th['border']}; border-radius: 4px; min-height: 40px; margin: 0 2px; }}
        QScrollBar::handle:vertical:hover {{ background: {th['muted']}; }}
        QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
        QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
        QScrollBar:horizontal {{ height: 0px; }}
        QDialog#Dialog, QWidget#Dialog, QMessageBox {{ background: {th['surface']}; }}
        #DialogTitle {{ font-size: 13pt; font-weight: 600; }}
        #Browser {{ background: {th['bg']}; }}
        #PageHeading {{ font-size: 20pt; font-weight: 700; padding: 4px 4px 0 4px; }}
        #Empty {{ color: {th['muted']}; font-size: 11pt; }}
        QListWidget#Grid {{ background: transparent; border: none; outline: none; }}
        QPushButton#Crumb, QPushButton#CrumbCur {{ background: transparent; border: none; border-radius: 14px;
            padding: 6px 10px; font-size: 11pt; color: {th['muted']}; }}
        QPushButton#CrumbCur {{ color: {th['text']}; font-weight: 600; }}
        QPushButton#Crumb:hover {{ background: {th['hover']}; color: {th['text']}; }}
        QPushButton#Crumb[drop="true"] {{ background: {th['soft']}; color: {th['accent']}; }}
        #CrumbSep {{ color: {th['muted']}; font-size: 12pt; padding: 0 2px; }}
        QMessageBox QLabel {{ min-width: 280px; }}
        QToolTip {{ background: {th['text']}; color: {th['bg']}; border: none; border-radius: 6px; padding: 5px 9px; }}
        """
        # общий «коэффициент скругления» интерфейса
        sheet = re.sub(r"(radius:\s*)(\d+)px", lambda m: f"{m.group(1)}{round(int(m.group(2)) * self.rf)}px", sheet)
        self.setStyleSheet(sheet)
        self._dark_titlebar(self.theme == "dark")

    def _update_color_icon(self):
        th = self.th
        col = getattr(self, "_last_color", th["danger"])
        self.color_btn.spec = ("glyph", "A", {"bar": col})
        self.color_btn.setIcon(make_icon(self.color_btn.spec, th))

    def _dark_titlebar(self, dark):
        if sys.platform != "win32":
            return
        try:
            import ctypes
            hwnd = int(self.winId())
            val = ctypes.c_int(1 if dark else 0)
            for attr in (20, 19):
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(val), ctypes.sizeof(val)) == 0:
                    break
        except Exception:
            pass

    def toggle_theme(self):
        self.theme_mode = "light" if self.theme == "dark" else "dark"
        self.settings.setValue("theme", self.theme_mode)
        self.theme = self.theme_mode
        self.apply_theme()

    def open_settings(self):
        SettingsDialog(self).exec()

    def _repolish(self, w):
        w.style().unpolish(w)
        w.style().polish(w)

    # ── Документ ──
    def doc(self):
        return self.editor.document()

    def setup_document(self):
        d = self.doc()
        f = QFont(DEFAULT_FONT)
        f.setPointSizeF(DEFAULT_SIZE)
        d.setDefaultFont(f)
        d.setDocumentMargin(56)
        ff = d.rootFrame().frameFormat()
        ff.setTopMargin(92)  # место под стеклянную панель
        d.rootFrame().setFrameFormat(ff)
        d.modificationChanged.connect(self.update_title)
        d.setUndoRedoEnabled(False)
        self.apply_spacing_all()
        d.setUndoRedoEnabled(True)
        d.setModified(False)
        self.undo_btn.setEnabled(False)
        self.redo_btn.setEnabled(False)

    def apply_spacing_all(self):
        cur = QTextCursor(self.doc())
        cur.select(QTextCursor.SelectionType.Document)
        bf = QTextBlockFormat()
        bf.setLineHeight(140, 1)  # ProportionalHeight
        bf.setBottomMargin(6)
        cur.mergeBlockFormat(bf)

    def apply_spacing_block(self):
        bf = QTextBlockFormat()
        bf.setLineHeight(140, 1)
        bf.setBottomMargin(6)
        self.editor.textCursor().mergeBlockFormat(bf)

    def maybe_save(self):
        if not self.doc().isModified():
            return True
        box = QMessageBox(self)
        box.setWindowTitle(APP_NAME)
        box.setIcon(QMessageBox.Icon.NoIcon)
        box.setText(f"Сохранить изменения в «{self.display_name()}»?")
        save = box.addButton("Сохранить", QMessageBox.ButtonRole.AcceptRole)
        discard = box.addButton("Не сохранять", QMessageBox.ButtonRole.DestructiveRole)
        box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        if clicked is save:
            return self.save()
        return clicked is discard

    def display_name(self):
        return Path(self.path).name if self.path else "Без названия"

    def update_title(self, *_):
        dirty = self.doc().isModified()
        name = self.display_name()
        self.title_label.setText(name + ("  •" if dirty else ""))
        if self.path:
            parent = Path(self.path).parent
            try:
                rel = parent.relative_to(self.browser.root)
                self.sub_label.setText(" › ".join(("Проекты",) + rel.parts))
            except ValueError:
                self.sub_label.setText(str(parent))
        else:
            self.sub_label.setText("Новый документ")
        self.setWindowTitle(f"{'• ' if dirty else ''}{name} — {APP_NAME}")
        self.format_chip.setText(self.fmt.upper() + (f" · {self.encoding.upper()}" if self.fmt == "txt" else ""))

    def new_document(self, ask=True):
        if ask and not self.maybe_save():
            return
        self.editor.setDocument(QTextDocument(self.editor))
        self.path, self.fmt, self.encoding, self.newline = None, "docx", "utf-8", "\r\n"
        self._target = None
        self.stickers.load(None)
        self.setup_document()
        self.update_title()
        self.update_counts()
        self.sync_toolbar()
        self.editor.setFocus()

    def open_dialog(self):
        if not self.maybe_save():
            return
        start = self.settings.value("last_dir", str(Path.home() / "Documents"))
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть", start,
            "Все документы (*.docx *.txt *.md *.markdown *.html *.htm *.log *.csv *.json *.xml *.ini);;"
            "Word (*.docx);;Текст (*.txt);;Markdown (*.md *.markdown);;Веб-страница (*.html *.htm);;Все файлы (*)")
        if path and self.open_path(path):
            self.stack.setCurrentWidget(self.editor_page)

    def open_path_checked(self, path):
        if self.maybe_save() and self.open_path(path):
            self.stack.setCurrentWidget(self.editor_page)

    def open_path(self, path):
        p = Path(path).resolve()
        ext = p.suffix.lower()
        if ext == ".doc":
            QMessageBox.information(self, APP_NAME, "Старый формат .doc (Word 97–2003) пока не поддерживается.\n"
                                                    "Folio открывает .docx, .txt, .md и .html.")
            return False
        if ext in {".rtf", ".odt", ".pdf"}:
            QMessageBox.information(self, APP_NAME, f"Формат {ext} не поддерживается.\n"
                                                    "Folio открывает .docx, .txt, .md и .html.")
            return False
        newdoc = QTextDocument(self.editor)
        try:
            if ext == ".docx":
                load_docx(str(p), newdoc)
                fmt = "docx"
            else:
                text, enc, nl = decode_text(p.read_bytes())
                if ext in MD_EXTS:
                    newdoc.setMarkdown(text)
                    fmt = "md"
                elif ext in HTML_EXTS:
                    newdoc.setHtml(text)
                    fmt = "html"
                else:
                    newdoc.setPlainText(text)
                    fmt = "txt"
                self.encoding, self.newline = enc, nl
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось открыть файл:\n{p.name}\n\n{e}")
            return False
        self.editor.setDocument(newdoc)
        self.path, self.fmt = str(p), fmt
        self._target = None
        self.stickers.load(str(p))
        if fmt == "docx":
            self.encoding = "utf-8"
        self.setup_document()
        self.settings.setValue("last_dir", str(p.parent))
        cur = self.editor.textCursor()
        cur.movePosition(QTextCursor.MoveOperation.Start)
        self.editor.setTextCursor(cur)
        self.update_title()
        self.update_counts()
        self.sync_toolbar()
        return True

    # ── Переходы между экранами ──
    def open_from_browser(self, path):
        if self.path and Path(self.path) == Path(path).resolve() and self.stack.currentWidget() is not self.editor_page:
            self.stack.setCurrentWidget(self.editor_page)  # уже открыт
            self.editor.setFocus()
            return
        if not self.maybe_save():
            return
        if self.open_path(path):
            self.stack.setCurrentWidget(self.editor_page)
            self.editor.setFocus()

    def go_home(self):
        if not self.maybe_save():
            return
        if self.findbar.isVisible():
            self.toggle_find(False)
        self.stickers.save()
        target = self.browser.cwd
        if self.path:
            parent = Path(self.path).parent
            try:
                parent.relative_to(self.browser.root)
                target = parent
            except ValueError:
                pass
        self.browser.navigate(target)
        if self.path:
            self.browser.select_path(Path(self.path))
        self.stack.setCurrentWidget(self.browser)
        self.browser.grid.setFocus()

    def new_shortcut(self):
        if self.stack.currentWidget() is self.browser:
            if self.browser.at_root():
                self.browser.new_folder()
            else:
                self.browser.new_doc()
        else:
            self.new_document()

    def path_moved(self, old, new):
        """Файл или папка переехали — если открытый документ был внутри, обновляем путь."""
        if not self.path:
            return
        cur, old, new = Path(self.path), Path(old), Path(new)
        try:
            rel = cur.relative_to(old)
        except ValueError:
            return
        self.path = str(new / rel) if str(rel) != "." else str(new)
        self.stickers.path = self.path
        self.update_title()

    def show_update_available(self, tag):
        self.browser.update_btn.setText(f"✨ Обновить до {tag.lstrip('v')}")
        self.browser.update_btn.show()

    def show_any_toast(self, text):
        if self.stack.currentWidget() is self.browser:
            self.show_browser_toast(text)
        else:
            self.show_toast(text)

    def prepare_quit(self):
        """Перед обновлением: сохранить стикеры, спросить про несохранённый документ."""
        if not self.maybe_save():
            return False
        self.stickers.save()
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("last_folder", str(self.browser.cwd))
        self.doc().setModified(False)
        return True

    def show_browser_toast(self, text):
        self.browser.toast.setText(text)
        self.browser.toast_timer.start(2500)

    # ── Ссылки в тексте ──
    def link_range(self, pos):
        """(start, end, href) ссылки вокруг позиции, или None."""
        block = self.tgt().document().findBlock(pos)
        frags = []
        it = block.begin()
        while not it.atEnd():
            f = it.fragment()
            it += 1
            if f.isValid():
                frags.append(f)
        for i, f in enumerate(frags):
            start, end = f.position(), f.position() + f.length()
            cf = f.charFormat()
            if cf.isAnchor() and start <= pos <= end and (pos < end or pos == end):
                href = cf.anchorHref()
                j = i
                while j > 0 and frags[j - 1].charFormat().isAnchor() and frags[j - 1].charFormat().anchorHref() == href:
                    j -= 1
                k = i
                while k + 1 < len(frags) and frags[k + 1].charFormat().isAnchor() \
                        and frags[k + 1].charFormat().anchorHref() == href:
                    k += 1
                return frags[j].position(), frags[k].position() + frags[k].length(), href
        return None

    def insert_link(self):
        cur = self.tgt().textCursor()
        rng = None
        if not cur.hasSelection():
            rng = self.link_range(cur.position())
            if rng is None and cur.position() > 0:
                rng = self.link_range(cur.position() - 1)
            if rng:
                cur.setPosition(rng[0])
                cur.setPosition(rng[1], QTextCursor.MoveMode.KeepAnchor)
        else:
            r2 = self.link_range(cur.selectionStart())
            if r2 and r2[1] >= cur.selectionEnd():
                rng = r2
        text = cur.selectedText().replace("\u2029", " ")
        href = rng[2] if rng else ""
        if not href:
            clip = QGuiApplication.clipboard().text().strip()
            if re.match(r"^https?://\S+$", text.strip()):
                href = text.strip()
            elif re.match(r"^https?://\S+$", clip):
                href = clip
        dlg = LinkDialog(self, "Изменить ссылку" if rng else "Вставить ссылку", "Готово",
                         "Текст (то, что видно в документе)", name=text, url=href, allow_remove=bool(rng))
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        if dlg.removed:
            self.remove_link(cur.selectionStart(), cur.selectionEnd())
            return
        name, url = dlg.values()
        if not url:
            return
        typed = dlg.name.text().strip() or url
        cur.beginEditBlock()
        if typed != text:
            fmt = strip_link_format(cur.charFormat() if cur.hasSelection() else self.tgt().currentCharFormat())
            cur.insertText(typed, fmt)
            start = cur.position() - len(typed)
            cur.setPosition(start)
            cur.setPosition(start + len(typed), QTextCursor.MoveMode.KeepAnchor)
        lf = QTextCharFormat()
        lf.setAnchor(True)
        lf.setAnchorHref(url)
        lf.setFontUnderline(True)
        lf.setForeground(QColor(LINK_COLOR))
        cur.mergeCharFormat(lf)
        end = cur.selectionEnd()
        cur.endEditBlock()
        cur.setPosition(end)
        self.tgt().setTextCursor(cur)
        self.tgt().setCurrentCharFormat(strip_link_format(self.tgt().currentCharFormat()))
        self.tgt().setFocus()

    def remove_link(self, start=None, end=None):
        if start is None:
            rng = self.link_range(self.tgt().textCursor().position())
            if not rng:
                return
            start, end, _ = rng
        cur = QTextCursor(self.tgt().document())
        cur.beginEditBlock()
        pos = start
        while pos < end:  # по фрагментам, чтобы сохранить жирный/курсив и т.п.
            c = QTextCursor(self.tgt().document())
            c.setPosition(pos)
            c.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
            c.setCharFormat(strip_link_format(c.charFormat()))
            pos += 1
        cur.endEditBlock()
        self.tgt().setFocus()

    def save(self):
        if not self.path:
            return self.save_as()
        return self.write_file(self.path, self.fmt)

    def save_as(self):
        if self.path:
            start_dir = str(Path(self.path).parent)
        elif not self.browser.at_root():
            start_dir = str(self.browser.cwd)
        else:
            start_dir = self.settings.value("last_dir", str(Path.home() / "Documents"))
        base = Path(self.path).stem if self.path else "Без названия"
        filters = {"Документ Word (*.docx)": "docx", "Текст (*.txt)": "txt",
                   "Markdown (*.md)": "md", "Веб-страница (*.html)": "html"}
        order = list(filters)
        current = next((f for f, k in filters.items() if k == self.fmt), order[0])
        path, chosen = QFileDialog.getSaveFileName(self, "Сохранить как", str(Path(start_dir) / f"{base}.{self.fmt}"),
                                                   ";;".join(order), current)
        if not path:
            return False
        fmt = filters.get(chosen, "docx")
        ext = Path(path).suffix.lower()
        by_ext = {".docx": "docx", ".txt": "txt", ".md": "md", ".markdown": "md", ".html": "html", ".htm": "html"}
        if ext in by_ext:
            fmt = by_ext[ext]
        else:
            path += "." + fmt
        if not self.write_file(path, fmt):
            return False
        self.path, self.fmt = path, fmt
        self.settings.setValue("last_dir", str(Path(path).parent))
        self.update_title()
        return True

    def write_file(self, path, fmt):
        try:
            if fmt == "docx":
                save_docx(self.doc(), path)
            else:
                if fmt == "txt":
                    text = self.doc().toPlainText()
                elif fmt == "md":
                    text = self.doc().toMarkdown()
                else:
                    text = self.doc().toHtml()
                text = text.replace(" ", "\n").replace("\r\n", "\n")
                if self.newline == "\r\n":
                    text = text.replace("\n", "\r\n")
                enc = self.encoding if fmt == "txt" else "utf-8"
                try:
                    data = text.encode(enc)
                except (UnicodeEncodeError, LookupError):
                    enc = "utf-8"
                    data = text.encode(enc)
                    self.encoding = enc
                Path(path).write_bytes(data)
        except PermissionError:
            QMessageBox.warning(self, APP_NAME, "Файл занят другой программой или нет прав на запись.\n"
                                                "Закрой его там или выбери «Сохранить как…».")
            return False
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось сохранить:\n{e}")
            return False
        self.doc().setModified(False)
        if self.stickers.path != str(path):
            self.stickers.set_path(str(path))
        else:
            self.stickers.save()
        self.show_toast("Сохранено ✓" if fmt != "txt" else "Сохранено как простой текст ✓")
        return True

    def export_pdf(self):
        self.export_as("pdf")

    def export_as(self, fmt, qdoc=None, src=None):
        """Сохраняет копию в другом формате. Открытый файл остаётся прежним."""
        qdoc = qdoc or self.doc()
        src = src or self.path
        ext = "." + fmt
        if src:
            folder = Path(src).parent
        elif not self.browser.at_root():
            folder = self.browser.cwd
        else:
            folder = Path(self.settings.value("last_dir", str(Path.home() / "Documents")))
        stem = Path(src).stem if src else "Без названия"
        label = next(l for f, l, _ in EXPORT_FORMATS if f == fmt)
        path, _ = QFileDialog.getSaveFileName(self, f"Экспорт — {label}", str(folder / f"{stem}{ext}"),
                                              f"{label} (*{ext})")
        if not path:
            return
        if not path.lower().endswith(ext):
            path += ext
        if self.path and Path(path).resolve() == Path(self.path).resolve() and fmt != self.fmt:
            QMessageBox.information(self, APP_NAME, "Выбери другое имя — иначе экспорт перезапишет открытый файл.")
            return
        try:
            export_document(qdoc, path, fmt)
        except PermissionError:
            QMessageBox.warning(self, APP_NAME, "Файл занят другой программой. Закрой его там или выбери другое имя.")
            return
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось экспортировать:\n{e}")
            return
        msg = f"Экспортировано: {Path(path).name} ✓"
        if self.stack.currentWidget() is self.browser:
            self.show_browser_toast(msg)
        else:
            self.show_toast(msg)

    def export_file(self, path, fmt):
        try:
            d = load_document(path)
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось открыть файл:\n{e}")
            return
        self.export_as(fmt, d, path)

    def print_doc(self):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dlg = QPrintDialog(printer, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._print_to(printer)

    def _print_to(self, printer):
        print_document(self.doc(), printer)

    def show_toast(self, text):
        self.toast.setText(text)
        self.toast_timer.start(2200)

    def update_counts(self):
        text = self.doc().toPlainText()
        words = len(re.findall(r"\w+", text))
        chars = len(text.replace("\n", "").replace(" ", ""))
        def plural(n, one, few, many):
            n10, n100 = n % 10, n % 100
            if n10 == 1 and n100 != 11:
                return one
            if 2 <= n10 <= 4 and not 12 <= n100 <= 14:
                return few
            return many
        self.count_label.setText(f"{words} {plural(words, 'слово', 'слова', 'слов')}  ·  "
                                 f"{chars} {plural(chars, 'символ', 'символа', 'символов')}")

    # ── Форматирование ──
    # ── Куда применять форматирование: основной текст или заметка в фокусе ──
    def tgt(self):
        t = getattr(self, "_target", None)
        if t is not None and t is not self.editor:
            try:
                if t.isVisible():
                    return t
            except RuntimeError:
                pass
            self._target = None
        return self.editor

    def set_target(self, ed):
        if ed is getattr(self, "_target", None):
            return
        self._target = ed
        self.sync_toolbar()
        for s_ in getattr(self, "stickers", None).items if hasattr(self, "stickers") else []:
            s_.update()

    def merge(self, cf):
        self.tgt().mergeCurrentCharFormat(cf)
        self.tgt().setFocus()

    def set_underline_style(self, style):
        cf = QTextCharFormat()
        cf.setUnderlineStyle(style)
        self.merge(cf)
        self.sync_toolbar()

    def set_underline_color(self, color):
        cf = QTextCharFormat()
        if color is None:
            cur = self.tgt().textCursor()
            if cur.hasSelection():
                self._clear_prop(QTextFormat.Property.TextUnderlineColor)
            return
        cf.setUnderlineColor(color)
        if self.tgt().currentCharFormat().underlineStyle() == QTextCharFormat.UnderlineStyle.NoUnderline:
            cf.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SingleUnderline)
        self.merge(cf)
        self.sync_toolbar()

    def set_bold(self, on):
        if self._updating:
            return
        cf = QTextCharFormat()
        cf.setFontWeight(700 if on else 400)
        self.merge(cf)

    def set_italic(self, on):
        if self._updating:
            return
        cf = QTextCharFormat()
        cf.setFontItalic(on)
        self.merge(cf)

    def set_underline(self, on):
        if self._updating:
            return
        cf = QTextCharFormat()
        cf.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SingleUnderline if on
                             else QTextCharFormat.UnderlineStyle.NoUnderline)
        self.merge(cf)

    def set_strike(self, on):
        if self._updating:
            return
        cf = QTextCharFormat()
        cf.setFontStrikeOut(on)
        self.merge(cf)

    def set_font_family(self, fam):
        cf = QTextCharFormat()
        cf.setFontFamilies([fam])
        self.merge(cf)

    def set_font_size(self, text):
        try:
            size = float(text.replace(",", "."))
        except ValueError:
            return
        if 1 <= size <= 400:
            cf = QTextCharFormat()
            cf.setFontPointSize(size)
            self.merge(cf)

    def set_text_color(self, color):
        cur = self.tgt().textCursor()
        if color is None:
            if cur.hasSelection():
                self._clear_prop(QTextFormat.Property.ForegroundBrush)
            else:
                f = self.tgt().currentCharFormat()
                f.clearForeground()
                self.tgt().setCurrentCharFormat(f)
        else:
            self._last_color = color.name()
            self._update_color_icon()
            cf = QTextCharFormat()
            cf.setForeground(color)
            self.merge(cf)
        self.tgt().setFocus()

    def set_highlight(self, color):
        cur = self.tgt().textCursor()
        if color is None:
            if cur.hasSelection():
                self._clear_prop(QTextFormat.Property.BackgroundBrush)
            else:
                f = self.tgt().currentCharFormat()
                f.clearBackground()
                self.tgt().setCurrentCharFormat(f)
        else:
            cf = QTextCharFormat()
            cf.setBackground(color)
            self.merge(cf)
        self.tgt().setFocus()

    def _clear_prop(self, prop):
        """Убирает свойство у каждого фрагмента выделения."""
        cur = self.tgt().textCursor()
        start, end = cur.selectionStart(), cur.selectionEnd()
        cur.beginEditBlock()
        pos = start
        while pos < end:
            c = QTextCursor(self.tgt().document())
            c.setPosition(pos)
            c.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
            f = c.charFormat()
            f.clearProperty(prop)
            c.setCharFormat(f)
            pos += 1
        cur.endEditBlock()

    def clear_formatting(self):
        cur = self.tgt().textCursor()
        if not cur.hasSelection():
            cur.select(QTextCursor.SelectionType.BlockUnderCursor)
        cf = QTextCharFormat()
        f = QFont(DEFAULT_FONT)
        f.setPointSizeF(DEFAULT_SIZE)
        cf.setFont(f)
        cur.setCharFormat(cf)
        self.tgt().setFocus()

    def set_block_style(self, level):
        cur = self.tgt().textCursor()
        cur.beginEditBlock()
        bf = QTextBlockFormat()
        bf.setHeadingLevel(level)
        cur.mergeBlockFormat(bf)
        sel = QTextCursor(cur)
        sel.movePosition(QTextCursor.MoveOperation.StartOfBlock)
        end = cur.selectionEnd() if cur.hasSelection() else cur.position()
        sel.setPosition(cur.selectionStart() if cur.hasSelection() else cur.position())
        sel.movePosition(QTextCursor.MoveOperation.StartOfBlock)
        sel.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        sel.movePosition(QTextCursor.MoveOperation.EndOfBlock, QTextCursor.MoveMode.KeepAnchor)
        cf = QTextCharFormat()
        if level:
            cf.setFontPointSize(HEADING_SIZES[level])
            cf.setFontWeight(700)
        else:
            cf.setFontPointSize(DEFAULT_SIZE)
            cf.setFontWeight(400)
        sel.mergeCharFormat(cf)
        sel.mergeBlockCharFormat(cf)
        cur.endEditBlock()
        self.tgt().mergeCurrentCharFormat(cf)
        self.sync_toolbar()
        self.tgt().setFocus()

    def set_align(self, flag):
        self.tgt().setAlignment(flag | Qt.AlignmentFlag.AlignAbsolute if flag == Qt.AlignmentFlag.AlignLeft else flag)
        self.sync_toolbar()
        self.tgt().setFocus()

    def toggle_list(self, kind, on):
        if self._updating:
            return
        cur = self.tgt().textCursor()
        cur.beginEditBlock()
        # сначала убираем выделенные абзацы из их списков (или меняем тип списка)
        for b in list(blocks_in(cur)):
            if b.textList() is not None or b.blockFormat().marker() != NO_MARKER:
                unlist_block(b)
        if on:
            make_list(cur, kind)
        cur.endEditBlock()
        self.sync_toolbar()
        self.tgt().setFocus()

    def change_list_indent(self, delta):
        cur = self.tgt().textCursor()
        lst = cur.currentList()
        if lst is None:
            return
        lf = QTextListFormat(lst.format())
        new = max(1, min(6, lf.indent() + delta))
        if new == lf.indent():
            return
        bullet = lf.style() in (QTextListFormat.Style.ListDisc, QTextListFormat.Style.ListCircle,
                                QTextListFormat.Style.ListSquare)
        block = cur.block()
        cur.beginEditBlock()
        # ищем соседний список такого же уровня выше
        target = None
        b = block.previous()
        while b.isValid() and b.textList() is not None:
            ind = b.textList().format().indent()
            if ind == new:
                target = b.textList()
                break
            if ind < new:
                break
            b = b.previous()
        lst.remove(block)
        if target is not None:
            target.add(block)
        else:
            lf.setIndent(new)
            if bullet:
                lf.setStyle([QTextListFormat.Style.ListDisc, QTextListFormat.Style.ListCircle,
                             QTextListFormat.Style.ListSquare][(new - 1) % 3])
            QTextCursor(block).createList(lf)
        cur.endEditBlock()

    def sync_toolbar(self):
        self._updating = True
        try:
            cf = self.tgt().currentCharFormat()
            cur = self.tgt().textCursor()
            self.bold_btn.setChecked(cf.fontWeight() >= 600)
            self.italic_btn.setChecked(cf.fontItalic())
            self.under_btn.setChecked(cf.underlineStyle() != QTextCharFormat.UnderlineStyle.NoUnderline
                                      and not cf.isAnchor())
            self.strike_btn.setChecked(cf.fontStrikeOut())
            fams = cf.fontFamilies() or []
            fam = fams[0] if fams else self.tgt().document().defaultFont().family()
            self.font_box.blockSignals(True)
            self.font_box.setCurrentFont(QFont(fam))
            self.font_box.blockSignals(False)
            size = cf.fontPointSize() or self.tgt().document().defaultFont().pointSizeF()
            self.size_box.setEditText(f"{size:g}")
            lvl = min(cur.blockFormat().headingLevel(), 3)
            self.style_box.setCurrentIndex(lvl)
            lst = cur.currentList()
            bullet = lst is not None and lst.format().style() in (
                QTextListFormat.Style.ListDisc, QTextListFormat.Style.ListCircle, QTextListFormat.Style.ListSquare)
            check = lst is not None and cur.blockFormat().marker() != NO_MARKER
            self.check_btn.setChecked(check)
            self.bullet_btn.setChecked(lst is not None and bullet and not check)
            self.number_btn.setChecked(lst is not None and not bullet)
            al = cur.blockFormat().alignment()
            if al & Qt.AlignmentFlag.AlignHCenter:
                active = Qt.AlignmentFlag.AlignHCenter
            elif al & Qt.AlignmentFlag.AlignJustify:
                active = Qt.AlignmentFlag.AlignJustify
            elif al & Qt.AlignmentFlag.AlignRight:
                active = Qt.AlignmentFlag.AlignRight
            else:
                active = Qt.AlignmentFlag.AlignLeft
            for flag, b in self.align_btns.items():
                b.setChecked(flag == active)
        finally:
            self._updating = False

    # ── Картинки и таблицы ──
    def insert_image(self):
        start = self.settings.value("last_img_dir", str(Path.home() / "Pictures"))
        path, _ = QFileDialog.getOpenFileName(self, "Картинка", start,
                                              "Изображения (*.png *.jpg *.jpeg *.gif *.bmp *.webp)")
        if path:
            self.settings.setValue("last_img_dir", str(Path(path).parent))
            self.insert_image_file(path)

    def insert_image_file(self, path):
        img = QImage(path)
        if img.isNull():
            QMessageBox.warning(self, APP_NAME, "Не получилось открыть картинку.")
            return
        self.insert_qimage(img)

    def insert_qimage(self, img):
        self._img_n = getattr(self, "_img_n", 0) + 1
        name = f"folio-img-{self._img_n}-{id(img)}"
        self.doc().addResource(QTextDocument.ResourceType.ImageResource, QUrl(name), img)
        f = QTextImageFormat()
        f.setName(name)
        w, h = img.width(), img.height()
        maxw = 640
        if w > maxw:
            h = h * maxw / w
            w = maxw
        f.setWidth(w)
        f.setHeight(h)
        self.editor.textCursor().insertImage(f)
        self.editor.setFocus()

    def insert_table(self):
        dlg = TableDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        tf = QTextTableFormat()
        tf.setBorder(1)
        tf.setBorderStyle(QTextFrameFormat.BorderStyle.BorderStyle_Solid)
        tf.setBorderBrush(QBrush(QColor("#C7CCD6")))
        tf.setBorderCollapse(True)
        tf.setCellPadding(6)
        tf.setCellSpacing(0)
        tf.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        self.editor.textCursor().insertTable(dlg.rows.value(), dlg.cols.value(), tf)
        self.editor.setFocus()

    def context_menu(self, pos):
        menu = self.editor.createStandardContextMenu(pos)
        href = self.editor.anchorAt(pos)
        if href:
            first = menu.actions()[0] if menu.actions() else None
            acts = [QAction("Открыть ссылку", menu), QAction("Изменить ссылку…", menu),
                    QAction("Копировать адрес", menu), QAction("Убрать ссылку", menu)]
            acts[0].triggered.connect(lambda: open_href(href, self))
            acts[1].triggered.connect(self.insert_link)
            acts[2].triggered.connect(lambda: QGuiApplication.clipboard().setText(href))
            acts[3].triggered.connect(lambda: self.remove_link())
            for a in acts:
                menu.insertAction(first, a)
            menu.insertSeparator(first)
        else:
            menu.addSeparator()
            menu.addAction("Вставить ссылку…", self.insert_link)
        menu.setWindowFlags(menu.windowFlags() | Qt.WindowType.FramelessWindowHint | Qt.WindowType.NoDropShadowWindowHint)
        menu.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        cur = self.editor.cursorForPosition(pos)
        if not self.editor.textCursor().hasSelection():
            self.editor.setTextCursor(cur)
        table = self.editor.textCursor().currentTable()
        if table is not None:
            cell = table.cellAt(self.editor.textCursor())
            r, c = cell.row(), cell.column()
            menu.addSeparator()
            menu.addAction("Строка ниже", lambda: table.insertRows(r + 1, 1))
            menu.addAction("Строка выше", lambda: table.insertRows(r, 1))
            menu.addAction("Столбец справа", lambda: table.insertColumns(c + 1, 1))
            menu.addAction("Столбец слева", lambda: table.insertColumns(c, 1))
            menu.addSeparator()
            menu.addAction("Удалить строку", lambda: table.removeRows(r, 1))
            menu.addAction("Удалить столбец", lambda: table.removeColumns(c, 1))
        menu.exec(self.editor.viewport().mapToGlobal(pos))

    # ── Поиск ──
    def toggle_find(self, show=None):
        show = (not self.findbar.isVisible()) if show is None else show
        self.findbar.setVisible(show)
        if show:
            sel = self.editor.textCursor().selectedText()
            if sel and " " not in sel:
                self.find_edit.setText(sel)
            self.find_edit.setFocus()
            self.find_edit.selectAll()
        else:
            self.editor.setFocus()

    def find_next(self, backward=False):
        text = self.find_edit.text()
        if not text:
            return False
        flags = QTextDocument.FindFlag.FindBackward if backward else QTextDocument.FindFlag(0)
        found = self.editor.find(text, flags)
        if not found:
            cur = self.editor.textCursor()
            cur.movePosition(QTextCursor.MoveOperation.End if backward else QTextCursor.MoveOperation.Start)
            self.editor.setTextCursor(cur)
            found = self.editor.find(text, flags)
        self.find_edit.setProperty("error", not found)
        self._repolish(self.find_edit)
        return found

    def replace_one(self):
        text = self.find_edit.text()
        if not text:
            return
        cur = self.editor.textCursor()
        if cur.hasSelection() and cur.selectedText().lower() == text.lower():
            cur.insertText(self.repl_edit.text())
        self.find_next()

    def replace_all(self):
        text = self.find_edit.text()
        if not text:
            return
        n = 0
        cur = QTextCursor(self.doc())
        cur.beginEditBlock()
        c = self.doc().find(text, 0)
        while not c.isNull():
            c.insertText(self.repl_edit.text())
            n += 1
            c = self.doc().find(text, c)
        cur.endEditBlock()
        self.show_toast(f"Заменено: {n}")

    # ── Окно ──
    def closeEvent(self, e):
        if self.maybe_save():
            self.stickers.save()
            self.settings.setValue("geometry", self.saveGeometry())
            self.settings.setValue("last_folder", str(self.browser.cwd))
            e.accept()
        else:
            e.ignore()


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Folio.Editor")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(APP_NAME)
    app.setStyle("Fusion")
    app.setWindowIcon(app_icon())
    try:
        load_user_fonts()
    except Exception:
        pass
    tr = QTranslator(app)
    if tr.load("qtbase_ru", QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)):
        app.installTranslator(tr)
    w = Folio()
    w.show()
    if len(sys.argv) > 1 and Path(sys.argv[1]).exists():
        if w.open_path(sys.argv[1]):
            w.stack.setCurrentWidget(w.editor_page)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
