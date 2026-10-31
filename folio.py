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
                               QCheckBox, QColorDialog, QLayout)
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtPrintSupport import QPrinter, QPrintDialog

APP_NAME = "Folio"
APP_VERSION = "1.3.0"
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
        try:
            pf = para.paragraph_format
            if pf.left_indent:
                bf.setLeftMargin(pf.left_indent.pt / 0.75)
            if pf.right_indent:
                bf.setRightMargin(pf.right_indent.pt / 0.75)
            if pf.first_line_indent:
                bf.setTextIndent(pf.first_line_indent.pt / 0.75)
        except Exception:
            pass
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
        pf = para.paragraph_format  # отступы абзаца (px при 96 dpi → pt)
        if bf.leftMargin() > 0.5 and not (lst is not None and marker != NO_MARKER):
            pf.left_indent = Pt(bf.leftMargin() * 0.75)
        if bf.rightMargin() > 0.5:
            pf.right_indent = Pt(bf.rightMargin() * 0.75)
        if abs(bf.textIndent()) > 0.5:
            pf.first_line_indent = Pt(bf.textIndent() * 0.75)

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


def paged_clone(qdoc):
    """Копия документа, разбитая на страницы A4 точно так, как она пойдёт в печать и PDF."""
    clone = qdoc.clone()
    clone.setDocumentMargin(0)
    ff = clone.rootFrame().frameFormat()
    ff.setMargin(0)
    clone.rootFrame().setFrameFormat(ff)
    w, h = page_px_size()
    clone.setPageSize(QSizeF(w, h))
    return clone


def print_document(qdoc, printer):
    printer.setPageLayout(QPageLayout(QPageSize(QPageSize.PageSizeId.A4), QPageLayout.Orientation.Portrait,
                                      QMarginsF(18, 18, 18, 18), QPageLayout.Unit.Millimeter))
    paged_clone(qdoc).print_(printer)


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
        if block.textList() is None:
            if bf.leftMargin() > 0.5:
                parts.append(f"\\li{round(bf.leftMargin() * 15)}")
            if abs(bf.textIndent()) > 0.5:
                parts.append(f"\\fi{round(bf.textIndent() * 15)}")
        if bf.rightMargin() > 0.5:
            parts.append(f"\\ri{round(bf.rightMargin() * 15)}")
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
    page_breaks_y = []   # [(номер страницы, y в документе)]
    guide_x = None

    def set_guide(self, x):
        self.guide_x = x
        self.viewport().update()

    def paintEvent(self, e):
        super().paintEvent(e)
        w = self.window()
        th = getattr(w, "th", None)
        if th is None:
            return
        p = QPainter(self.viewport())
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        scroll = self.verticalScrollBar().value()
        vw = self.viewport().width()
        if self.page_breaks_y and getattr(w, "show_breaks", True):
            col = QColor(th["muted"])
            col.setAlpha(150)
            f = QFont(getattr(w, "ui_font", "Segoe UI"))
            f.setPixelSize(10)
            f.setBold(True)
            p.setFont(f)
            fm = QFontMetrics(f)
            for n, y_doc in self.page_breaks_y:
                y = y_doc - scroll - 5
                if y < -20 or y > self.viewport().height() + 20:
                    continue
                pen = QPen(col, 1, Qt.PenStyle.DashLine)
                pen.setDashPattern([5, 4])
                p.setPen(pen)
                p.drawLine(QPointF(10, y), QPointF(vw - 10, y))
                label = f"стр. {n}"
                lw = fm.horizontalAdvance(label) + 16
                r = QRectF(vw - lw - 14, y - 9, lw, 18)
                p.setPen(Qt.PenStyle.NoPen)
                p.setBrush(QColor(th["hover"]))
                p.drawRoundedRect(r, 9, 9)
                p.setPen(QColor(th["muted"]))
                p.drawText(r, Qt.AlignmentFlag.AlignCenter, label)
        if self.guide_x is not None:
            pen = QPen(QColor(th["accent"]), 1, Qt.PenStyle.DashLine)
            p.setPen(pen)
            p.drawLine(QPointF(self.guide_x, 0), QPointF(self.guide_x, self.viewport().height()))

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


def ask_copy_or_link(parent, n, names=None):
    box = QMessageBox(parent)
    box.setWindowTitle(APP_NAME)
    box.setIcon(QMessageBox.Icon.NoIcon)
    what = f"«{names[0]}»" if names and n == 1 else ("файл" if n == 1 else f"{n} {plural(n, 'элемент', 'элемента', 'элементов')}")
    box.setText(f"Как добавить {what}?")
    box.setInformativeText(
        "📄 Скопировать — Folio сделает свою копию в «Документы\\Folio». "
        "Оригинал на компьютере останется как был, правки в Folio его не трогают.\n\n"
        "🔗 Ссылкой — ничего не копируется, файл или папка остаются на своём месте. "
        "Folio просто показывает их здесь, и все правки идут прямо в оригинал. "
        "Если оригинал удалить или переместить — ссылка перестанет работать.")
    copy = box.addButton("Скопировать", QMessageBox.ButtonRole.AcceptRole)
    link = box.addButton("Ссылкой", QMessageBox.ButtonRole.ActionRole)
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
            url = read_url_file(p)
            target = Path(QUrl(url).toLocalFile()) if QUrl(url).isLocalFile() else None
            linked_dir = target is not None and target.is_dir()
            items.append(dict(kind="link", path=str(p), name=p.stem, url=url, mtime=mtime, linked_dir=linked_dir,
                              color=folder_color(target) if linked_dir else None, broken=bool(target) and not target.exists()))
        elif p.suffix.lower() in DOC_EXTS:
            items.append(dict(kind="doc", path=str(p), name=p.stem, ext=p.suffix.lower(), mtime=mtime))
        else:
            items.append(dict(kind="file", path=str(p), name=p.stem, ext=p.suffix.lower(), mtime=mtime))
    order = {"folder": 0, "doc": 1, "link": 2, "file": 3}
    items.sort(key=lambda it: (0 if it.get("linked_dir") else order[it["kind"]], it["name"].casefold()))
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
        elif kind == "link" and it.get("linked_dir"):
            p.fillPath(pp, QColor(it["color"]))
            pm = cached_svg("folder", "#FFFFFF", 46)
            p.drawPixmap(int(prev.center().x() - 23), int(prev.center().y() - 30), pm)
            badge = QRectF(prev.right() - 34, prev.y() + 8, 26, 26)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(255, 255, 255, 235))
            p.drawEllipse(badge)
            p.drawPixmap(int(badge.x() + 4), int(badge.y() + 4), cached_svg("link", it["color"], 18))
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
            if it.get("broken"):
                sub = "⚠ оригинал не найден"
            elif it.get("linked_dir"):
                sub = "Папка-ссылка"
            else:
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
        if item and not item.isSelected():
            it = item.data(Qt.ItemDataRole.UserRole)
            if it["kind"] == "folder" or it.get("linked_dir"):
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
        if row >= 0:
            it = self.item(row).data(Qt.ItemDataRole.UserRole)
            target = Path(QUrl(it["url"]).toLocalFile()) if it.get("linked_dir") else Path(it["path"])
        else:
            target = self.page.cwd
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
        self.ext = None  # папка-ссылка, в которой мы сейчас (оригинал вне «Документы\\Folio»)
        self.setObjectName("Browser")

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        bar = QWidget()
        bar.setObjectName("AppBar")
        h = QHBoxLayout(bar)
        h.setContentsMargins(16, 12, 16, 4)
        h.setSpacing(0)
        # слева: путь + круглая кнопка «+»
        left = GlassBar(win)
        lh = QHBoxLayout(left)
        lh.setContentsMargins(8, 5, 6, 5)
        lh.setSpacing(2)
        logo = QLabel()
        logo.setFixedSize(32, 32)
        logo.setPixmap(app_icon().pixmap(28, 28))
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lh.addWidget(logo)
        lh.addSpacing(4)
        self.crumbs = QHBoxLayout()
        self.crumbs.setSpacing(0)
        lh.addLayout(self.crumbs)
        lh.addSpacing(6)
        self.create_btn = QToolButton()
        self.create_btn.setObjectName("AddBtn")
        self.create_btn.setFixedSize(36, 36)
        self.create_btn.setIconSize(QSize(18, 18))
        self.create_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.create_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.create_btn.clicked.connect(self.show_create_menu)
        lh.addWidget(self.create_btn)
        h.addWidget(left)
        h.addStretch()
        # справа: сообщения, обновление и настройки
        right = GlassBar(win)
        rh = QHBoxLayout(right)
        rh.setContentsMargins(12, 5, 6, 5)
        rh.setSpacing(2)
        self.toast = QLabel("")
        self.toast.setObjectName("Toast")
        self.toast_timer = QTimer(self)
        self.toast_timer.setSingleShot(True)
        self.toast_timer.timeout.connect(lambda: self.toast.setText(""))
        rh.addWidget(self.toast)
        self.update_btn = QPushButton("Обновить")
        self.update_btn.setObjectName("Ghost")
        self.update_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.update_btn.hide()
        self.update_btn.clicked.connect(lambda: win.updater.ask_install())
        rh.addWidget(self.update_btn)
        rh.addWidget(win._btn("Открыть файл с компьютера  (Ctrl+O)", ("svg", "open"), win.open_dialog))
        self.theme_btn = win._btn("Тёмная / светлая тема", ("svg", "moon"), win.toggle_theme)
        rh.addWidget(self.theme_btn)
        rh.addWidget(win._btn("Настройки: цвета, фон окна, шрифты", ("svg", "tune"), win.open_settings))
        h.addWidget(right)
        v.addWidget(bar)

        body = QWidget()
        body.setObjectName("EditorArea")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(9, 10, 9, 0)
        self.heading = QLabel()  # заголовок теперь в «пути» сверху
        self.grid = Grid(self)
        self.grid.setItemDelegate(CardDelegate(win))
        self.grid.itemDoubleClicked.connect(self.open_item)
        self.grid.customContextMenuRequested.connect(self.context_menu)
        bl.addWidget(self.grid, 1)
        self.empty = QLabel()
        self.empty.setObjectName("EmptyPill")
        self.empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty.setWordWrap(True)
        self.empty.setMaximumWidth(560)
        bl.addWidget(self.empty, 1, Qt.AlignmentFlag.AlignCenter)
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

    def inside(self, path, base):
        try:
            Path(path).relative_to(base)
            return True
        except ValueError:
            return False

    def navigate(self, path, ext_name=None):
        """Внутри «Документы\\Folio» или внутри папки-ссылки (оригинал лежит в другом месте)."""
        path = Path(path)
        if not path.exists() or not path.is_dir():
            path = self.root
        if ext_name is not None:
            self.ext = (path, ext_name)
        if self.inside(path, self.root):
            self.ext = None
        elif not (getattr(self, "ext", None) and self.inside(path, self.ext[0])):
            path, self.ext = self.root, None
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
        names = {}
        ext = getattr(self, "ext", None)
        base = ext[0] if ext else self.root
        if ext:
            chain.append(ext[0])
            names[ext[0]] = "🔗 " + ext[1]
        if self.cwd != base:
            acc = base
            for part in self.cwd.relative_to(base).parts:
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
                b = CrumbButton("Проекты" if c == self.root else names.get(c, c.name), c, self, current=(c == self.cwd))
            self.crumbs.addWidget(b)
        self.heading.setText("Мои проекты" if self.at_root() else self.cwd.name)
        self.create_btn.setToolTip("Новый проект" if self.at_root() else "Создать: документ, папку, ссылку")

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
            mode = ask_copy_or_link(self, len(local), [p.name for p in local])
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
        elif it["kind"] == "link" and it.get("linked_dir"):
            self.navigate(QUrl(it["url"]).toLocalFile(), ext_name=it["name"])
        elif it["kind"] == "link":
            if it.get("broken"):
                QMessageBox.information(self, APP_NAME, "Оригинал этой ссылки не найден — его удалили или переместили.\n"
                                                        "Правый клик → «Изменить ссылку», чтобы указать новое место.")
            elif it.get("url"):
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
        self.glass = QCheckBox("Размытие фона под панелями (стекло, как в iPhone)")
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
        self.sl_board = slider_row("Ширина доски заметок", 180, 520, float(self.s.value("board_width", 300)),
                                   "board_width", 1, " px")
        self.sl_tint = slider_row("Плотность панелей", 5, 100, panel_alpha(self.win) * 100,
                                  "glass_tint", 0.01, "%")
        self.sl_blur = slider_row("Размытие под панелями", 2, 30, float(self.s.value("glass_blur", 12)), "glass_blur", 1)
        self.sl_tbs = slider_row("Размер панели текста", 70, 150, float(self.s.value("toolbar_scale", 1.0)) * 100,
                                 "toolbar_scale", 0.01, "%")
        self.sl_tbf = slider_row("Шрифт на панели текста", 7, 13, float(self.s.value("toolbar_font", 10)),
                                 "toolbar_font", 1, " pt")
        wall_row = QHBoxLayout()
        wl = QLabel("Фон окна")
        wl.setFixedWidth(170)
        wall_row.addWidget(wl)
        cur_wall = self.s.value("wallpaper", "gradient")
        for key, label in (("none", "Без фона"), ("gradient", "Градиент"), ("photo", "Своё фото…")):
            b = QPushButton(label)
            b.setObjectName("Seg")
            b.setCheckable(True)
            b.setAutoExclusive(True)
            b.setChecked(key == cur_wall or (key == "photo" and cur_wall not in ("none", "gradient")))
            b.clicked.connect(lambda _=False, k=key: self.win.set_wallpaper_mode(k))
            wall_row.addWidget(b)
        wall_row.addStretch()
        lay.addLayout(wall_row)
        for key, label in (("rulers", "Линейка отступов (тяни бегунки)"),
                           ("page_breaks", "Показывать, где заканчиваются страницы (для печати и PDF)"),
                           ("links", "Линии от стикеров к строкам текста")):
            cb = QCheckBox(label)
            cb.setChecked(self.s.value(key, "true") in (True, "true"))
            cb.toggled.connect(lambda on, k=key: (self.s.setValue(k, "true" if on else "false"), self.win.apply_theme()))
            lay.addWidget(cb)
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
                     "snap", "inspector_auto", "recent_colors", "toolbar_scale", "toolbar_font",
                     "rulers", "page_breaks", "links", "wallpaper", "board_width", "doc_panel_open")

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
        for k in ("toolbar_scale", "toolbar_font", "links", "wallpaper", "board_width", "rulers", "page_breaks", "ui_round", "ui_font", "page_width", "glass_tint", "glass_blur", "snap", "inspector_auto",
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
    def __init__(self, parent, value=None, title=None):
        super().__init__(parent, "Дедлайн", "Готово")
        from PySide6.QtWidgets import QDateTimeEdit
        self.name = self.field("Что нужно сделать", title or "", "Например: выложить серию") if title is not None else None
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
        if value and title is None:
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

    def title_text(self):
        return self.name.text().strip() if self.name is not None else ""


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
    HEADER = 16   # полоска сверху с ручкой для перетаскивания
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
        self._link_drag = False
        self.hovered = False
        self.setMouseTracking(True)
        self.setAttribute(Qt.WidgetAttribute.WA_Hover)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.setMinimumSize(90, 50)
        self.chip = QPushButton(self)
        self.chip.setObjectName("DeadlineChip")
        self.chip.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chip.setToolTip("Клик — отметить выполненным")
        self.chip.clicked.connect(self.toggle_done)
        self.chip.hide()

    # ── переопределяется в наследниках ──
    def bg(self):
        c = QColor(self.d.get("color", STICKY_COLORS[0]))
        c.setAlphaF(max(0.0, min(1.0, self.d.get("opacity", 0.95))))
        return c

    def fg(self):
        return "#1A1D24" if luminance(self.d.get("color", "#FFE97A")) > 0.35 else "#F2F4F8"

    def content_rect(self):
        m = self.MARGIN
        r = QRect(m + 10, m + self.HEADER, self.width() - 2 * m - 20, self.height() - 2 * m - self.HEADER - 8)
        if self.d.get("shape") == "ellipse":  # в овале текст держим ближе к центру
            dx, dy = int(self.width() * 0.11), int(self.height() * 0.08)
            r.adjust(dx, dy - 6, -dx, -dy)
        elif self.d.get("shape") == "pill":
            dx = int(min(self.width(), self.height()) * 0.12)
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
        # ручка для перетаскивания — кружок в правом верхнем углу
        gc = self.grip_center()
        dark_bg = luminance(self.d.get("color", "#FFFFFF")) < 0.35 if self.d.get("type") != "photo" else True
        gcol = QColor(255, 255, 255) if dark_bg else QColor(40, 44, 56)
        gcol.setAlpha(150 if (self.hovered or self.is_selected()) else 70)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(gcol)
        p.drawEllipse(gc, 8, 8)
        # кружок для новой стрелки — на правом краю, виден при наведении
        if self.hovered or self._link_drag:
            acc = QColor(self.layer.win.th["accent"])
            p.setPen(QPen(acc, 2))
            p.setBrush(acc if self._link_drag else QColor("#FFFFFF"))
            p.drawEllipse(self.knob_center(), 5.5, 5.5)

    def grip_center(self):
        return QPointF(self.width() - self.MARGIN - 13, self.MARGIN + 13)

    def in_move_grip(self, pos):
        c = self.grip_center()
        return (pos.x() - c.x()) ** 2 + (pos.y() - c.y()) ** 2 <= 11 ** 2

    def knob_center(self):
        return QPointF(self.width() - self.MARGIN, self.height() / 2)

    def in_knob(self, pos):
        c = self.knob_center()
        return (pos.x() - c.x()) ** 2 + (pos.y() - c.y()) ** 2 <= 10 ** 2

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self.delete()
            return
        super().keyPressEvent(e)

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
        self.setFocus()
        self.layer.select(self)
        if e.button() != Qt.MouseButton.LeftButton:
            return
        pos = e.position().toPoint()
        g = e.globalPosition().toPoint()
        if self.in_knob(pos):
            self._link_drag = True
            self._moved = True
            return
        if self.in_grip(pos):
            self._resize = (g, self.d["w"], self.d["h"])
        elif self.in_move_grip(pos) or self.draggable_at(pos):
            self._drag = (g, self.d["x"], self.d["y"])

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        g = e.globalPosition().toPoint()
        if self._link_drag:
            ov = self.layer.win.links
            ov.set_temp(ov.local(self, self.knob_center().toPoint()), ov.mapFromGlobal(g))
            self.layer.highlight_drop(ov.mapFromGlobal(g), exclude=self.d.get("id"))
            return
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
            self.d["y"] = max(0, min(y, self.layer.max_y(self.d["h"])))
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
            if self.in_knob(pos):
                self.setCursor(Qt.CursorShape.CrossCursor)
                self.setToolTip("Потяни к строке текста или к другому стикеру — появится стрелка")
            elif self.in_grip(pos):
                self.setCursor(Qt.CursorShape.SizeFDiagCursor)
                self.setToolTip("")
            elif self.in_move_grip(pos):
                self.setCursor(Qt.CursorShape.SizeAllCursor)
                self.setToolTip("Перетащить")
            elif self.draggable_at(pos):
                self.setCursor(Qt.CursorShape.SizeAllCursor)
            else:
                self.setCursor(Qt.CursorShape.ArrowCursor)

    def mouseReleaseEvent(self, e):
        if self._link_drag:
            self._link_drag = False
            ov = self.layer.win.links
            ov.set_temp(None)
            self.layer.highlight_drop(None)
            self.layer.finish_new_connector(self, ov.mapFromGlobal(e.globalPosition().toPoint()))
            self.update()
            return
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
        self.update()

    def leaveEvent(self, e):
        self.hovered = False
        self.update()

    def resizeEvent(self, e):
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
        self.layer.add_connector(from_sticker=self)

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


def draw_arrow_head(p, tip, ctrl, w, col, kind):
    """Наконечник стрелки в точке tip; ctrl — точка, откуда линия к ней приходит."""
    import math
    from PySide6.QtGui import QPolygonF
    if kind == "none":
        return
    ang = math.atan2(tip.y() - ctrl.y(), tip.x() - ctrl.x())
    size = 9 + w * 2.2
    p.setPen(QPen(col, w, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))

    def at(a, r):
        return QPointF(tip.x() + r * math.cos(a), tip.y() + r * math.sin(a))
    if kind == "dot":
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(col)
        r = 2.5 + w * 0.9
        p.drawEllipse(tip, r, r)
    elif kind == "arrow":
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


class Connector:
    """Одна стрелка на все случаи: соединяет стикеры, строки текста или свободные точки.
    Рисуется поверх доски и страницы, поэтому спокойно переходит с одной на другую."""

    def __init__(self, layer, d):
        self.layer = layer
        self.d = d
        if not d.get("id"):
            import uuid
            d["id"] = uuid.uuid4().hex[:10]
        d["type"] = "arrow"
        if "heads" in d and "head_end" not in d:  # старый формат
            h = d.pop("heads")
            d["head_end"] = "none" if h == "none" else "arrow"
            d["head_start"] = "arrow" if h == "both" else "none"
        if isinstance(d.get("dash"), bool):
            d["line"] = "dash" if d.pop("dash") else "solid"
        d.setdefault("bend", [0, 0])
        doc = layer.editor.document()
        for k in (1, 2):
            if d.get(f"t{k}") is not None:
                c = QTextCursor(doc)
                c.setPosition(max(0, min(int(d[f"t{k}"]), doc.characterCount() - 1)))
                layer.anchors[(d["id"], k)] = c

    # ── концы ──
    def sticker(self, k):
        return self.layer.by_id(self.d.get(f"a{k}"))

    def cursor(self, k):
        return self.layer.anchors.get((self.d["id"], k))

    def kind(self, k):
        if self.sticker(k) is not None:
            return "sticker"
        if self.cursor(k) is not None:
            return "text"
        return "free"

    def free_point(self, k):
        ov, vp = self.layer.win.links, self.layer.viewport
        return QPointF(ov.local(vp, QPoint(int(self.d.get(f"x{k}", 0)),
                                         int(self.d.get(f"y{k}", 0) - self.layer.scroll()))))

    def set_free(self, k, ov_pt):
        ov_pt = QPointF(ov_pt)
        ov, vp = self.layer.win.links, self.layer.viewport
        local = vp.mapFrom(ov.parentWidget(), (ov_pt + QPointF(ov.pos())).toPoint())
        self.d[f"x{k}"] = local.x()
        self.d[f"y{k}"] = local.y() + self.layer.scroll()

    def rough(self, k):
        s = self.sticker(k)
        if s is not None:
            return self.layer.win.links.sticker_rect(s).center()
        c = self.cursor(k)
        if c is not None:
            return self.layer.win.links.text_point(c)[0]
        return self.free_point(k)

    def anchor(self, k, toward):
        """(точка, направление выхода, обрезано ли)"""
        s = self.sticker(k)
        if s is not None:
            r = self.layer.win.links.sticker_rect(s)
            c = r.center()
            dx, dy = toward.x() - c.x(), toward.y() - c.y()
            other = 2 if k == 1 else 1
            if self.kind(other) == "text" and dx > 0:  # к тексту всегда выходим с правого края
                return QPointF(r.right() + 1, c.y()), QPointF(1, 0), False
            if abs(dx) / max(1.0, r.width()) >= abs(dy) / max(1.0, r.height()):
                return (QPointF(r.right() + 1, c.y()), QPointF(1, 0), False) if dx >= 0 else \
                    (QPointF(r.left() - 1, c.y()), QPointF(-1, 0), False)
            return (QPointF(c.x(), r.bottom() + 1), QPointF(0, 1), False) if dy >= 0 else \
                (QPointF(c.x(), r.top() - 1), QPointF(0, -1), False)
        cur = self.cursor(k)
        if cur is not None:
            pt, clipped = self.layer.win.links.text_point(cur)
            return pt, QPointF(-1, 0), clipped
        return self.free_point(k), QPointF(0, 0), False

    def geometry(self):
        """Путь стрелки в координатах слоя + точки для наконечников и ручки изгиба."""
        r1, r2 = self.rough(1), self.rough(2)
        p1, n1, cl1 = self.anchor(1, r2)
        p2, n2, cl2 = self.anchor(2, r1)
        bx, by = self.d.get("bend", [0, 0])
        shape = self.d.get("shape", "curve")
        path = QPainterPath(p1)
        if shape == "straight":
            path.lineTo(p2)
            handle = (p1 + p2) / 2
        elif shape == "elbow":
            mx = (p1.x() + p2.x()) / 2 + bx
            path.lineTo(QPointF(mx, p1.y()))
            path.lineTo(QPointF(mx, p2.y()))
            path.lineTo(p2)
            handle = QPointF(mx, (p1.y() + p2.y()) / 2)
        else:
            import math
            dist = math.hypot(p2.x() - p1.x(), p2.y() - p1.y())
            k = float(self.d.get("curv", 1.0)) * max(40.0, dist * 0.45)
            c1 = p1 + n1 * k + QPointF(bx, by)
            c2 = p2 + n2 * k + QPointF(bx, by)
            path.cubicTo(c1, c2, p2)
            handle = path.pointAtPercent(0.5)
        t1 = path.pointAtPercent(0.04) if path.length() > 1 else p2
        t2 = path.pointAtPercent(0.96) if path.length() > 1 else p1
        return dict(p1=p1, p2=p2, path=path, handle=handle, t1=t1, t2=t2, clipped=cl1 or cl2)

    # ── действия ──
    def update(self):
        self.layer.win.links.refresh()

    def set(self, key, val):
        if val is None:
            self.d.pop(key, None)
        else:
            self.d[key] = val
        self.update()
        self.layer.mark()

    def attach(self, k, ov_pt):
        """Прицепить конец k к тому, что под точкой: стикеру, строке текста или оставить свободным."""
        ov_pt = QPointF(ov_pt)
        other = self.d.get(f"a{2 if k == 1 else 1}")
        s = self.layer.sticker_at_ov(ov_pt, exclude=other)
        self.layer.anchors.pop((self.d["id"], k), None)
        self.d.pop(f"t{k}", None)
        self.d[f"a{k}"] = None
        if s is not None:
            self.d[f"a{k}"] = s.d["id"]
            return
        ed = self.layer.editor
        vp_pt = ed.viewport().mapFrom(self.layer.win.links.parentWidget(),
                                      (ov_pt + QPointF(self.layer.win.links.pos())).toPoint())
        if ed.viewport().rect().contains(vp_pt):
            c = QTextCursor(ed.document())
            c.setPosition(ed.cursorForPosition(vp_pt).position())
            self.layer.anchors[(self.d["id"], k)] = c
            self.d[f"t{k}"] = c.position()
            return
        self.set_free(k, ov_pt)

    def freeze(self, k):
        g = self.geometry()
        self.set_free(k, g["p1"] if k == 1 else g["p2"])
        self.d[f"a{k}"] = None
        self.layer.anchors.pop((self.d["id"], k), None)
        self.d.pop(f"t{k}", None)

    def detach(self):
        for k in (1, 2):
            if self.kind(k) != "free":
                self.freeze(k)
        self.update()
        self.layer.mark()

    def reverse(self):
        d = self.d
        A = self.layer.anchors
        i = d["id"]
        c1, c2 = A.pop((i, 1), None), A.pop((i, 2), None)
        if c2 is not None:
            A[(i, 1)] = c2
        if c1 is not None:
            A[(i, 2)] = c1
        for a, b in (("a1", "a2"), ("x1", "x2"), ("y1", "y2"), ("t1", "t2"), ("head_start", "head_end")):
            d[a], d[b] = d.get(b), d.get(a)
        for key in list(d):
            if d[key] is None and key in ("t1", "t2", "head_start", "head_end"):
                d.pop(key)
        self.update()
        self.layer.mark()

    def straighten(self):
        self.set("bend", [0, 0])

    def delete(self):
        self.layer.remove_connector(self)

    def snippet(self, k):
        s = self.sticker(k)
        if s is not None:
            t = {"note": "заметка", "photo": "фото", "link": "ссылка", "deadline": "дедлайн"}.get(s.d.get("type"), "стикер")
            return t
        c = self.cursor(k)
        if c is not None:
            txt = c.block().text().strip()
            return "строка «" + ((txt[:26] + "…") if len(txt) > 27 else (txt or "пустая")) + "»"
        return "свободный конец"


class ConnectorHit(QWidget):
    """Невидимый слой для мыши: ловит клики только по линиям стрелок и их ручкам."""
    def __init__(self, win, parent):
        super().__init__(parent)
        self.win = win
        self._drag = None
        self.hover = None
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.menu)

    def hit(self, pt):
        L = self.win.stickers
        sel = L.selected if isinstance(L.selected, Connector) else None
        geos = self.win.links.geos
        if sel is not None and sel.d["id"] in geos:
            g = geos[sel.d["id"]]
            for part, q in (("p1", g["p1"]), ("p2", g["p2"]), ("h", g["handle"])):
                if part == "h" and sel.d.get("shape") == "straight":
                    continue
                if (q.x() - pt.x()) ** 2 + (q.y() - pt.y()) ** 2 <= 11 ** 2:
                    return sel, part
        from PySide6.QtGui import QPainterPathStroker
        st = QPainterPathStroker()
        for c in reversed(L.conns):
            g = geos.get(c.d["id"])
            if g is None:
                continue
            st.setWidth(max(12.0, float(c.d.get("width", 2)) + 10))
            if st.createStroke(g["path"]).contains(pt):
                return c, "line"
        return None, None

    def mousePressEvent(self, e):
        pt = e.position()
        c, part = self.hit(pt)
        if c is None:
            e.ignore()
            return
        self.setFocus()
        self.win.stickers.select(c)
        if e.button() == Qt.MouseButton.LeftButton:
            g = self.win.links.geos.get(c.d["id"])
            self._drag = (c, part, pt, list(c.d.get("bend", [0, 0])), g["p1"], g["p2"], False)

    def mouseMoveEvent(self, e):
        pt = e.position()
        if not self._drag:
            c, part = self.hit(pt)
            self.setCursor(Qt.CursorShape.SizeAllCursor if part in ("p1", "p2", "h") else
                           Qt.CursorShape.PointingHandCursor if part == "line" else Qt.CursorShape.ArrowCursor)
            if c is not self.hover:
                self.hover = c
                self.win.links.refresh()
            return
        c, part, start, bend, p1, p2, moved = self._drag
        dx, dy = pt.x() - start.x(), pt.y() - start.y()
        if abs(dx) + abs(dy) > 2:
            self._drag = (c, part, start, bend, p1, p2, True)
        L = self.win.stickers
        if part in ("p1", "p2"):
            k = 1 if part == "p1" else 2
            c.d[f"a{k}"] = None
            L.anchors.pop((c.d["id"], k), None)
            c.d.pop(f"t{k}", None)
            c.set_free(k, pt)
            L.highlight_drop(pt, exclude=c.d.get(f"a{2 if k == 1 else 1}"))
        elif part == "h":
            c.d["bend"] = [bend[0] + dx, bend[1] + (dy if c.d.get("shape") != "elbow" else 0)]
        else:  # двигаем свободные концы вместе
            if c.kind(1) == "free":
                c.set_free(1, p1 + QPointF(dx, dy))
            if c.kind(2) == "free":
                c.set_free(2, p2 + QPointF(dx, dy))
        c.update()

    def mouseReleaseEvent(self, e):
        if not self._drag:
            return
        c, part, start, bend, p1, p2, moved = self._drag
        self._drag = None
        if part in ("p1", "p2"):
            self.win.stickers.highlight_drop(None)
            c.attach(1 if part == "p1" else 2, e.position())
        c.update()
        if moved:
            self.win.stickers.mark()
            if self.win.inspector.item is c:
                self.win.inspector.rebuild()

    def mouseDoubleClickEvent(self, e):
        c, part = self.hit(e.position())
        if c is not None and part == "h":
            c.straighten()

    def keyPressEvent(self, e):
        sel = self.win.stickers.selected
        if isinstance(sel, Connector) and e.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            sel.delete()
            return
        super().keyPressEvent(e)

    def wheelEvent(self, e):
        QApplication.sendEvent(self.win.editor.verticalScrollBar(), e)

    def menu(self, pos):
        c, part = self.hit(QPointF(pos))
        if c is None:
            return
        self.win.stickers.select(c)
        m = styled_menu(self)
        m.addAction("Развернуть", c.reverse)
        m.addAction("Выпрямить изгиб", c.straighten)
        m.addAction("Отцепить концы", c.detach)
        m.addSeparator()
        m.addAction("Удалить стрелку", c.delete)
        m.exec(self.mapToGlobal(pos))

    def set_region(self, paths, handles):
        from PySide6.QtGui import QPainterPathStroker, QRegion
        region = QRegion()
        st = QPainterPathStroker()
        for path, w in paths:
            st.setWidth(max(14.0, w + 12))
            region = region.united(QRegion(st.createStroke(path).toFillPolygon().toPolygon()))
        for q in handles:
            region = region.united(QRegion(int(q.x() - 12), int(q.y() - 12), 24, 24, QRegion.RegionType.Ellipse))
        if region.isEmpty():
            region = QRegion(-10, -10, 1, 1)
        if region != self.mask():
            self.setMask(region)


class DeadlineSticker(Sticker):
    """Дедлайн — отдельная карточка: иконка часов, что сделать и когда."""
    def __init__(self, layer, data):
        super().__init__(layer, data)
        self.setMinimumSize(170, 64)
        self.setToolTip("Клик по часам — отметить выполненным")
        self.refresh_style()

    def draggable_at(self, pos):
        return True

    def when(self):
        try:
            return datetime.fromisoformat(self.d.get("deadline", ""))
        except (TypeError, ValueError):
            return None

    def circle(self):
        m = self.MARGIN
        r = QRectF(self.rect()).adjusted(m, m, -m, -m)
        return QRectF(r.left() + 12, r.center().y() - 21, 42, 42)

    def paint_body(self, p, r):
        super().paint_body(p, r)
        th = self.layer.win.th
        fg = QColor(self.fg())
        d = self.when()
        done = self.d.get("done", False)
        state = deadline_color(d, done) if d else None
        col = QColor(state or th["accent"])
        circle = self.circle()
        p.setPen(Qt.PenStyle.NoPen)
        c2 = QColor(col)
        c2.setAlpha(40 if luminance(self.d.get("color", "#FFF")) > 0.35 else 70)
        p.setBrush(c2)
        p.drawEllipse(circle)
        icon = "check" if done else "clock"
        p.drawPixmap(int(circle.center().x() - 13), int(circle.center().y() - 13),
                     svg_pixmap(SVG.get(icon, SVG["clock"]), col.name(), size=26, stroke=2.1))
        tx = circle.right() + 12
        w = int(r.right() - tx - 28)
        f = QFont(self.layer.win.ui_font)
        f.setPointSizeF(10.5)
        f.setWeight(QFont.Weight.DemiBold)
        f.setStrikeOut(done)
        p.setFont(f)
        p.setPen(fg)
        fm = QFontMetrics(f)
        p.drawText(QRectF(tx, r.center().y() - fm.height() + 1, w, fm.height()), 0,
                   fm.elidedText(self.d.get("title") or "Дедлайн", Qt.TextElideMode.ElideRight, w))
        f2 = QFont(self.layer.win.ui_font)
        f2.setPointSizeF(8.5)
        f2.setWeight(QFont.Weight.DemiBold)
        p.setFont(f2)
        p.setPen(col if state else QColor(fg.red(), fg.green(), fg.blue(), 150))
        if d is None:
            sub = "без даты"
        elif done:
            sub = "сделано"
        else:
            left = (d - datetime.now()).total_seconds()
            sub = deadline_text(d) + ("  ·  просрочено" if left < 0 else "  ·  меньше суток" if left < 86400 else "")
        p.drawText(QRectF(tx, r.center().y() + 3, w, 18), 0, QFontMetrics(f2).elidedText(sub, Qt.TextElideMode.ElideRight, w))

    def mouseReleaseEvent(self, e):
        was_drag = self._moved
        super().mouseReleaseEvent(e)
        if not was_drag and e.button() == Qt.MouseButton.LeftButton and self.circle().contains(e.position()):
            self.toggle_done()
            self.update()
            if self.layer.win.inspector.item is self:
                self.layer.win.inspector.rebuild()

    def update_chip(self):
        self.chip.hide()

    def edit_deadline(self):
        dlg = DeadlineDialog(self.layer.win, self.d.get("deadline"), title=self.d.get("title", ""))
        if dlg.exec() != QDialog.DialogCode.Accepted or dlg.removed:
            return False
        self.d["deadline"] = dlg.value()
        self.d["title"] = dlg.title_text()
        self.d["done"] = False
        self.update()
        self.layer.mark()
        return True


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
        self.body.document().setIndentWidth(18)
        if data.get("html"):
            self.body.setHtml(data["html"])
        self.body.textChanged.connect(self._text_changed)
        self.body.cursorPositionChanged.connect(lambda: layer.win.sync_toolbar() if layer.win.tgt() is self.body else None)
        self.body.currentCharFormatChanged.connect(lambda _: layer.win.sync_toolbar() if layer.win.tgt() is self.body else None)
        self.refresh_style()

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
            top = self.MARGIN + 30
        right = r.right() + 4
        if top < self.MARGIN + 26:  # первая строка не должна залезать под ручку
            right = min(right, int(self.grip_center().x() - 12))
        self.body.setGeometry(r.left() - 4, top, max(20, right - r.left() + 4), r.bottom() - top + 6)

    def draggable_at(self, pos):
        return pos.y() <= self.MARGIN + 10 and not self.chip.geometry().contains(pos)

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
        p.restore()


class LinkSticker(Sticker):
    def __init__(self, layer, data):
        super().__init__(layer, data)
        self.setMinimumSize(170, 70)
        self.setToolTip("Двойной клик — открыть ссылку")
        self.refresh_style()

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
        w = int(r.right() - tx - 28)
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

    def mouseDoubleClickEvent(self, e):
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
    """Стикеры живут на доске слева, прокручиваются вместе с текстом и хранятся в скрытом файле рядом с документом.
    Стрелки (Connector) — отдельный список, рисуются поверх доски и страницы."""
    def __init__(self, win, editor, viewport):
        super().__init__(win)
        self.win = win
        self.editor = editor
        self.viewport = viewport
        self.items = []     # стикеры
        self.conns = []     # стрелки
        self.anchors = {}   # (id стрелки, конец) → QTextCursor — едет вместе с текстом при правках
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
        vw = self.viewport.width()
        if vw >= 120:  # доска уже разложена — подгоняем стикер под её ширину
            if s.d["w"] > vw + 2 * Sticker.MARGIN - 8:
                ratio = s.photo_ratio() if hasattr(s, "photo_ratio") else None
                s.d["w"] = vw + 2 * Sticker.MARGIN - 8
                if ratio:
                    s.d["h"] = round((s.d["w"] - 2 * Sticker.MARGIN) / ratio) + 2 * Sticker.MARGIN
            s.d["x"] = max(-Sticker.MARGIN, min(s.d["x"], vw - s.d["w"] + Sticker.MARGIN))
        s.setGeometry(int(s.d["x"]), int(s.d["y"] - self.scroll()), int(s.d["w"]), int(s.d["h"]))
        self.win.links.refresh()

    def max_y(self, h):
        sb = self.editor.verticalScrollBar()
        return max(0, sb.maximum() + self.viewport.height() - h)

    def relayout(self):
        for s in self.items:
            self.place(s)
        self.win.links.refresh()

    def select(self, item):
        if item is self.selected:
            return
        old = self.selected
        self.selected = item
        for it in (old, item):
            if it is not None:
                try:
                    it.update()
                except RuntimeError:
                    pass
        self.win.links.refresh()
        if item is None:
            self.win.inspector.hide_panel()
        else:
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
            if o is st:
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
        self.mark()

    def by_id(self, sid):
        if not sid:
            return None
        for s in self.items:
            if s.d.get("id") == sid:
                return s
        return None

    def sticker_at_ov(self, ov_pt, exclude=None):
        """Стикер под точкой слоя стрелок (только в видимой части доски)."""
        ov = self.win.links
        ov_pt = QPointF(ov_pt)
        vp_rect = QRectF(QPointF(ov.local(self.viewport, QPoint(0, 0))), QSizeF(self.viewport.size()))
        if not vp_rect.contains(ov_pt):
            return None
        for s in reversed(self.items):
            if s.d.get("id") == exclude:
                continue
            if ov.sticker_rect(s).adjusted(-6, -6, 6, 6).contains(ov_pt):
                return s
        return None

    def highlight_drop(self, ov_pt, exclude=None):
        target = self.sticker_at_ov(ov_pt, exclude) if ov_pt is not None else None
        for s in self.items:
            flag = s is target
            if getattr(s, "drop_target", False) != flag:
                s.drop_target = flag
                s.update()

    # ── стрелки ──
    def create_connector(self, d):
        c = Connector(self, d)
        self.conns.append(c)
        return c

    def add_connector(self, from_sticker=None):
        col = self.win.th["accent"]
        if from_sticker is not None:
            r = from_sticker.doc_rect()
            d = dict(type="arrow", a1=from_sticker.d["id"], a2=None, x2=r.center().x(), y2=r.bottom() + 110,
                     color=col, width=2, head_start="dot", head_end="arrow")
        else:
            vw = self.viewport.width() or 300
            y = self.scroll() + 200
            d = dict(type="arrow", a1=None, a2=None, x1=vw * 0.2, y1=y, x2=vw * 0.8, y2=y + 40,
                     color=col, width=2, head_start="none", head_end="arrow")
        c = self.create_connector(d)
        self.win.links.hit.setFocus()
        self.select(c)
        self.mark()
        return c

    def finish_new_connector(self, sticker, ov_pt):
        """Потянули кружок с края стикера и отпустили — создаём стрелку к тому, что под курсором."""
        ov_pt = QPointF(ov_pt)
        start = self.win.links.sticker_rect(sticker)
        if start.adjusted(-4, -4, 4, 4).contains(ov_pt):
            return None  # отпустили на самом стикере — ничего не делаем
        d = dict(type="arrow", a1=sticker.d["id"], a2=None, color=self.win.th["accent"], width=2,
                 head_start="dot", head_end="dot")
        c = self.create_connector(d)
        c.attach(2, ov_pt)
        if c.kind(2) == "free":
            c.d["head_end"] = "arrow"
        self.select(c)
        self.mark()
        if c.kind(2) == "text":
            self.win.show_toast("Связано со строкой ✓")
        return c

    def remove_connector(self, c):
        if c in self.conns:
            self.conns.remove(c)
        for k in (1, 2):
            self.anchors.pop((c.d["id"], k), None)
        if self.selected is c:
            self.select(None)
        self.win.links.refresh()
        self.mark()

    # ── загрузка / сохранение ──
    def clear(self):
        self.selected = None
        self.anchors = {}
        self.conns = []
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
        self.load_data(read_sidecar(doc_path)["stickers"] if doc_path else [], doc_path)

    def dump(self):
        """Копия всех стикеров и стрелок (для переключения вкладок)."""
        for c in self.conns:
            for k in (1, 2):
                cur = self.anchors.get((c.d["id"], k))
                if cur is not None:
                    c.d[f"t{k}"] = cur.position()
        return json.loads(json.dumps([s.d for s in self.items] + [c.d for c in self.conns], ensure_ascii=False))

    def load_data(self, data, doc_path):
        self.save_timer.stop()
        self.clear()
        self.path = doc_path
        if data:
            for d in data:
                if d.get("type") == "arrow":
                    continue
                try:
                    self.create(d)
                except Exception:
                    continue
            for d in data:
                if d.get("type") != "arrow":
                    continue
                try:
                    self.create_connector(d)
                except Exception:
                    continue
            # старые «связи с текстом» (1.1.0) превращаем в обычные стрелки
            for s in self.items:
                if s.d.get("link_pos") is not None:
                    self.create_connector(dict(type="arrow", a1=s.d["id"], t2=s.d.pop("link_pos"),
                                               color=s.d.pop("link_c", None) or self.win.th["accent"],
                                               width=2, head_start="dot", head_end="dot"))
                s.d.pop("link_c", None)
            self.relayout()

    def create(self, d):
        if d.get("type") == "arrow":
            return self.create_connector(d)
        cls = {"note": NoteSticker, "photo": PhotoSticker, "link": LinkSticker,
               "deadline": DeadlineSticker}.get(d.get("type", "note"), NoteSticker)
        s = cls(self, d)
        s.show()
        self.items.append(s)
        self.place(s)
        self.viewport.update()
        return s

    def mark(self):
        self.save_timer.start()

    def save(self):
        for c in self.conns:
            for k in (1, 2):
                cur = self.anchors.get((c.d["id"], k))
                if cur is not None:
                    c.d[f"t{k}"] = cur.position()
        if not self.path:
            return
        sc = sidecar_path(self.path)
        if not self.items and not self.conns:
            if sc.exists():
                try:
                    sc.unlink()
                except OSError:
                    pass
            return
        try:
            _write_hidden(sc, json.dumps({"version": 2, "stickers": [s.d for s in self.items] +
                                          [c.d for c in self.conns]}, ensure_ascii=False))
        except OSError:
            pass

    def set_path(self, path):
        self.path = path
        self.save()

    def refresh_styles(self):
        for s in self.items:
            s.refresh_style()
        self.win.links.refresh()

    # ── добавление ──
    def spot(self, w, h, pos=None):
        vw = self.viewport.width()
        if pos is not None:
            return max(0, min(pos.x() - w // 2, vw - w)), max(0, pos.y() + self.scroll() - 20)
        n = len(self.items) % 5
        return max(0, min(6 + 14 * n, vw - w)), self.scroll() + 52 + 26 * n

    def note_size(self):
        vw = self.viewport.width() or 300
        return max(150, min(240, vw - 8)), 120

    def add_note(self, kind="note"):
        w, h = self.note_size()
        x, y = self.spot(w, h)
        d = dict(type="note", x=x, y=y, w=w, h=h, color=STICKY_COLORS[0], opacity=0.95, html="")
        if kind == "todo":
            d["color"] = STICKY_COLORS[2]
            d["h"] = 140
        s = self.create(d)
        s.body.setFocus()
        if kind == "todo":
            s.add_checklist()
        self.select(s)
        self.mark()
        return s

    def add_deadline(self):
        dlg = DeadlineDialog(self.win, None, title="")
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        w = self.note_size()[0]
        x, y = self.spot(w, 80)
        s = self.create(dict(type="deadline", x=x, y=y, w=w, h=80, color="#FFFFFF", opacity=0.97,
                             title=dlg.title_text(), deadline=dlg.value(), done=False))
        self.select(s)
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
        w = self.note_size()[0]
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
        w, h = self.note_size()[0], 80
        x, y = self.spot(w, h)
        s = self.create(dict(type="link", x=x, y=y, w=w, h=h, color="#FFFFFF", opacity=0.97, title=name, url=url))
        self.mark()
        return s

    def remove(self, s):
        sid = s.d.get("id")
        for c in list(self.conns):  # стрелки, прицепленные к стикеру, уходят вместе с ним
            if sid in (c.d.get("a1"), c.d.get("a2")):
                self.remove_connector(c)
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
        self.viewport.update()
        self.mark()


# ───────────────────────────── Стеклянная панель ─────────────────────────────
class FlowLayout(QLayout):
    """Раскладка «как текст»: элементы в строку, а если не влезают — переносятся на следующую."""
    def __init__(self, parent=None, hspacing=2, vspacing=4):
        super().__init__(parent)
        self.items = []
        self.hs, self.vs = hspacing, vspacing

    def addItem(self, item):
        self.items.append(item)

    def addSpacing(self, w):
        sp = QWidget()
        sp.setFixedSize(w, 1)
        sp.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.addWidget(sp)

    def count(self):
        return len(self.items)

    def itemAt(self, i):
        return self.items[i] if 0 <= i < len(self.items) else None

    def takeAt(self, i):
        return self.items.pop(i) if 0 <= i < len(self.items) else None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, w):
        return self._do(QRect(0, 0, w, 0), True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def one_row_width(self):
        m = self.contentsMargins()
        ws = [it.sizeHint().width() for it in self.items if not it.isEmpty()]
        return sum(ws) + self.hs * max(0, len(ws) - 1) + m.left() + m.right()

    def minimumSize(self):
        s = QSize()
        for it in self.items:
            s = s.expandedTo(it.minimumSize())
        m = self.contentsMargins()
        return s + QSize(m.left() + m.right(), m.top() + m.bottom())

    def _do(self, rect, test):
        m = self.contentsMargins()
        r = rect.adjusted(m.left(), m.top(), -m.right(), -m.bottom())
        rows, row, x = [], [], r.x()
        for it in self.items:
            if it.isEmpty():
                continue
            w = it.sizeHint().width()
            if row and x + w > r.right() + 1:
                rows.append(row)
                row, x = [], r.x()
            row.append(it)
            x += w + self.hs
        if row:
            rows.append(row)
        # строка не должна начинаться с разделителя
        y = r.y()
        for row in rows:
            while row and row[0].widget() is not None and row[0].widget().objectName() in ("Sep", ""):
                if row[0].widget().objectName() == "Sep" or row[0].widget().maximumHeight() <= 1:
                    if not test:
                        row[0].setGeometry(QRect(-100, -100, 0, 0))
                    row = row[1:]
                else:
                    break
            if not row:
                continue
            h = max(it.sizeHint().height() for it in row)
            total = sum(it.sizeHint().width() for it in row) + self.hs * (len(row) - 1)
            x = r.x() + max(0, (r.width() - total) // 2)  # строки по центру
            if not test:
                for it in row:
                    sh = it.sizeHint()
                    it.setGeometry(QRect(QPoint(x, y + (h - sh.height()) // 2), sh))
                    x += sh.width() + self.hs
            y += h + self.vs
        return y - self.vs - rect.y() + m.bottom() if rows else m.top() + m.bottom()


def page_px_size():
    """Область печати A4 (поля 18 мм) в пикселях при 96 dpi — так же, как её раскладывает печать Qt."""
    mm = 96 / 25.4
    return (210 - 36) * mm, (297 - 36) * mm


def compute_page_breaks(qdoc):
    """Позиции в тексте, с которых начинается каждая следующая страница при печати/PDF."""
    clone = paged_clone(qdoc)
    w, h = page_px_size()
    n = clone.pageCount()
    lay = clone.documentLayout()
    out = []
    for k in range(1, n):
        pos = lay.hitTest(QPointF(4, k * h + 3), Qt.HitTestAccuracy.FuzzyHit)
        if pos >= 0:
            out.append((k + 1, pos))
    return out


class RulerH(QWidget):
    """Горизонтальная линейка с бегунками отступов, как в Word."""
    H = 26

    def __init__(self, win):
        super().__init__()
        self.win = win
        self.setFixedHeight(self.H)
        self.setMouseTracking(True)
        self._drag = None
        self.setToolTip("Тяни бегунки: ▽ — первая строка, △ — отступ абзаца слева, △ справа — отступ справа.\n"
                        "Двойной клик — сбросить отступы. Alt — без привязки к шагу.")

    def ppcm(self):
        return self.logicalDpiX() / 2.54

    def geom(self):
        """x начала и конца области текста в координатах линейки."""
        ed = self.win.editor
        vp = ed.viewport()
        left = vp.mapTo(self.window(), QPoint(0, 0)).x() - self.mapTo(self.window(), QPoint(0, 0)).x()
        m = ed.document().documentMargin()
        return left + m, left + vp.width() - m

    def block_fmt(self):
        return self.win.editor.textCursor().blockFormat()

    def markers(self):
        x0, x1 = self.geom()
        bf = self.block_fmt()
        lst = self.win.editor.textCursor().currentList()
        list_ind = lst.format().indent() * self.win.editor.document().indentWidth() if lst is not None else 0
        left = x0 + bf.leftMargin() + list_ind
        return {"first": left + bf.textIndent(), "left": left, "right": x1 - bf.rightMargin()}

    def paintEvent(self, e):
        th = self.win.th
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        x0, x1 = self.geom()
        ppcm = self.ppcm()
        h = self.height()
        rr = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        paint_pane(self, p, rr, rr.height() / 2, self.win)
        track = QColor(th["text"])
        track.setAlpha(28)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(track)
        p.drawRoundedRect(QRectF(x0, h / 2 + 1, x1 - x0, 3), 1.5, 1.5)
        f = QFont(self.win.ui_font)
        f.setPixelSize(9)
        p.setFont(f)
        muted = QColor(th["muted"])
        i = 0
        while True:
            x = x0 + i * ppcm / 4
            if x > x1 + 0.5:
                break
            if i % 4 == 0:
                if i:
                    p.setPen(muted)
                    p.drawText(QRectF(x - 12, 4, 24, 12), Qt.AlignmentFlag.AlignCenter, str(i // 4))
            else:
                p.setPen(QPen(muted, 1))
                tick = 5 if i % 2 == 0 else 3
                p.drawLine(QPointF(x, h / 2 - tick / 2 + 3), QPointF(x, h / 2 + tick / 2 + 3))
            i += 1
        acc = QColor(th["accent"])
        mk = self.markers()
        from PySide6.QtGui import QPolygonF
        def tri(x, down):
            y = 3 if down else h - 3
            d = 7 if down else -7
            return QPolygonF([QPointF(x - 5.5, y), QPointF(x + 5.5, y), QPointF(x, y + d)])
        p.setPen(QPen(QColor("#FFFFFF"), 1))
        p.setBrush(acc)
        p.drawPolygon(tri(mk["first"], True))
        p.drawPolygon(tri(mk["left"], False))
        p.drawPolygon(tri(mk["right"], False))

    def hit(self, x, y):
        mk = self.markers()
        top = y < self.height() / 2
        order = (["first", "left", "right"] if top else ["left", "right", "first"])
        for k in order:
            if abs(mk[k] - x) <= 7:
                return k
        return None

    def mousePressEvent(self, e):
        k = self.hit(e.position().x(), e.position().y())
        if k:
            self._drag = k
            self.win.editor.guide_x = None

    def mouseMoveEvent(self, e):
        x = e.position().x()
        if not self._drag:
            self.setCursor(Qt.CursorShape.SizeHorCursor if self.hit(x, e.position().y()) else Qt.CursorShape.ArrowCursor)
            return
        x0, x1 = self.geom()
        step = self.ppcm() / 4
        snap = not (e.modifiers() & Qt.KeyboardModifier.AltModifier)
        if snap:
            x = x0 + round((x - x0) / step) * step
        x = max(x0, min(x1 - 20, x)) if self._drag != "right" else max(x0 + 40, min(x1, x))
        bf = self.block_fmt()
        lst = self.win.editor.textCursor().currentList()
        list_ind = lst.format().indent() * self.win.editor.document().indentWidth() if lst is not None else 0
        nf = QTextBlockFormat()
        if self._drag == "left":
            first_off = bf.textIndent()
            nf.setLeftMargin(max(0.0, x - x0 - list_ind))
            nf.setTextIndent(first_off)
        elif self._drag == "first":
            nf.setTextIndent(x - (x0 + bf.leftMargin() + list_ind))
        else:
            nf.setRightMargin(max(0.0, x1 - x))
        self.win.editor.textCursor().mergeBlockFormat(nf)
        self.win.editor.set_guide(x - self.geom()[0] + self.win.editor.document().documentMargin())
        self.update()

    def mouseReleaseEvent(self, e):
        if self._drag:
            self._drag = None
            self.win.editor.set_guide(None)

    def mouseDoubleClickEvent(self, e):
        nf = QTextBlockFormat()
        nf.setLeftMargin(0)
        nf.setRightMargin(0)
        nf.setTextIndent(0)
        self.win.editor.textCursor().mergeBlockFormat(nf)
        self.update()


class RulerV(QWidget):
    """Вертикальная линейка: сантиметры и границы страниц."""
    W = 24

    def __init__(self, win):
        super().__init__()
        self.win = win
        self.setFixedWidth(self.W)

    def paintEvent(self, e):
        th = self.win.th
        ed = self.win.editor
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        vp_top = ed.viewport().mapTo(self.window(), QPoint(0, 0)).y() - self.mapTo(self.window(), QPoint(0, 0)).y()
        scroll = ed.verticalScrollBar().value()
        top = ed.document().rootFrame().frameFormat().topMargin()
        ppcm = self.logicalDpiY() / 2.54
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(th["hover"]))
        p.drawRoundedRect(QRectF(3, vp_top, self.width() - 6, ed.viewport().height()), 5, 5)
        f = QFont(self.win.ui_font)
        f.setPixelSize(9)
        p.setFont(f)
        muted = QColor(th["muted"])
        p.setClipRect(QRect(0, vp_top, self.width(), ed.viewport().height()))
        # сантиметры от начала каждой страницы
        starts = [0.0] + [y for _, y in ed.page_breaks_y]
        for si, start in enumerate(starts):
            end = starts[si + 1] if si + 1 < len(starts) else ed.document().size().height()
            base = (top if si == 0 else start)
            i = 1
            while True:
                y_doc = base + i * ppcm / 2
                if y_doc >= end - 4:
                    break
                y = vp_top + y_doc - scroll
                if y > vp_top + ed.viewport().height() + 20:
                    break
                if y >= vp_top - 20:
                    if i % 2 == 0:
                        p.setPen(muted)
                        p.drawText(QRectF(0, y - 6, self.width(), 12), Qt.AlignmentFlag.AlignCenter, str(i // 2))
                    else:
                        p.setPen(QPen(muted, 1))
                        p.drawLine(QPointF(self.width() / 2 - 2, y), QPointF(self.width() / 2 + 2, y))
                i += 1
        # границы страниц
        acc = QColor(th["accent"])
        for n, y_doc in ed.page_breaks_y:
            y = vp_top + y_doc - scroll
            p.setPen(QPen(acc, 2))
            p.drawLine(QPointF(4, y), QPointF(self.width() - 4, y))


# ───────────────────────────── Фон и полупрозрачные панели ─────────────────────────────
_WALL_CACHE = {}


def blur_image(img, k):
    w, h = img.width(), img.height()
    for kk in (max(2, k), max(2, k // 2)):  # уменьшить-увеличить дважды = мягкое размытие
        small = img.scaled(max(1, w // kk), max(1, h // kk), Qt.AspectRatioMode.IgnoreAspectRatio,
                           Qt.TransformationMode.SmoothTransformation)
        img = small.scaled(w, h, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
    return img


def gradient_wallpaper(accent, dark):
    """Мягкий фон из цветных пятен в тон цвету интерфейса."""
    key = (accent, dark)
    if key in _WALL_CACHE:
        return _WALL_CACHE[key]
    from PySide6.QtGui import QRadialGradient
    W, H = 1600, 1000
    img = QImage(W, H, QImage.Format.Format_RGB32)
    img.fill(QColor("#0E1120" if dark else "#E7ECF6"))
    p = QPainter(img)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    blobs = [(0.16, 0.78, 0.55, accent), (0.04, 0.18, 0.42, mix(accent, "#12A594", 0.6)),
             (0.88, 0.88, 0.5, mix(accent, "#F76B15", 0.55)), (0.82, 0.12, 0.46, mix(accent, "#D6409F", 0.6)),
             (0.52, 0.46, 0.34, mix(accent, "#FFFFFF", 0.55))]
    for cx, cy, r, col in blobs:
        c = QColor(mix(col, "#000000", 0.35)) if dark else QColor(col)
        g = QRadialGradient(cx * W, cy * H, r * W)
        c0 = QColor(c)
        c0.setAlpha(150 if dark else 175)
        c1 = QColor(c)
        c1.setAlpha(0)
        g.setColorAt(0, c0)
        g.setColorAt(1, c1)
        p.fillRect(img.rect(), g)
    p.end()
    img = blur_image(img, 10)
    _WALL_CACHE[key] = img
    return img


class WallRoot(QWidget):
    """Корень окна: рисует фон (картинку или цвет) и держит его размытую копию для «стекла»."""
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.src = None
        self.key = None
        self.pm = None
        self.blur = None
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(60)
        self.timer.timeout.connect(self.rebuild)

    def set_source(self, img, key):
        if key == self.key and self.pm is not None:
            return
        self.key = key
        self.src = img
        self.rebuild()

    def resizeEvent(self, e):
        self.timer.start()

    def rebuild(self):
        if self.src is None or self.src.isNull() or self.width() < 10:
            self.pm = self.blur = None
            self.update()
            return
        dpr = self.devicePixelRatioF()
        size = QSize(int(self.width() * dpr), int(self.height() * dpr))
        scaled = self.src.scaled(size, Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                                 Qt.TransformationMode.SmoothTransformation)
        x, y = (scaled.width() - size.width()) // 2, (scaled.height() - size.height()) // 2
        scaled = scaled.copy(x, y, size.width(), size.height())
        if self.win.theme == "dark" and self.win.settings.value("wallpaper", "gradient") not in ("gradient", "none"):
            p = QPainter(scaled)  # фото в тёмной теме чуть притемняем
            p.fillRect(scaled.rect(), QColor(0, 0, 0, 90))
            p.end()
        self.pm = QPixmap.fromImage(scaled)
        self.pm.setDevicePixelRatio(dpr)
        k = int(self.win.settings.value("glass_blur", 12))
        half = scaled.scaled(max(1, scaled.width() // 2), max(1, scaled.height() // 2),
                             Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        bl = blur_image(half, max(2, k)).scaled(scaled.size(), Qt.AspectRatioMode.IgnoreAspectRatio,
                                                Qt.TransformationMode.SmoothTransformation)
        self.blur = QPixmap.fromImage(bl)
        self.blur.setDevicePixelRatio(dpr)
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(self.win.th["bg"]))
        if self.pm is not None:
            p.drawPixmap(self.rect(), self.pm)


def panel_alpha(win):
    try:
        return max(0.05, min(1.0, float(win.settings.value("glass_tint", 0.55))))
    except (TypeError, ValueError):
        return 0.55


def paint_pane(widget, p, r, radius, win):
    """Полупрозрачная панель: размытый фон под ней + лёгкая заливка + тонкая светлая кромка."""
    th = win.th
    dark = win.theme == "dark"
    path = QPainterPath()
    path.addRoundedRect(r, radius, radius)
    wall = getattr(win, "wall", None)
    alpha = panel_alpha(win)
    glass = win.settings.value("glass", "true") in (True, "true")
    if wall is not None and not wall.isAncestorOf(widget):
        wall = None  # панель в отдельном окне (диалог) — рисуем без фона
    if wall is not None and wall.pm is not None and alpha < 0.999:
        src = wall.blur if (glass and wall.blur is not None) else None
        if src is not None:
            off = widget.mapTo(wall, QPoint(0, 0))
            p.save()
            p.setClipPath(path)
            p.drawPixmap(-off.x(), -off.y(), src)
            p.restore()
        tint = QColor(th["surface"])
        tint.setAlphaF(alpha)
    else:
        tint = QColor(th["surface"])
    p.fillPath(path, tint)
    if glass and wall is not None and wall.pm is not None:
        from PySide6.QtGui import QLinearGradient
        g = QLinearGradient(0, r.top(), 0, r.bottom())
        g.setColorAt(0, QColor(255, 255, 255, 26 if dark else 90))
        g.setColorAt(0.6, QColor(255, 255, 255, 0))
        p.fillPath(path, g)
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.setPen(QPen(QColor(255, 255, 255, 40 if dark else 150) if wall is not None and wall.pm is not None
                  else QColor(th["border"]), 1))
    p.drawPath(path)


class PageTrack(QWidget):
    """Вертикальная полоса прокрутки по страницам: отметки — где начинаются страницы."""
    W = 16
    PAD = 6

    def __init__(self, win):
        super().__init__()
        self.win = win
        self.setFixedWidth(self.W)
        self.setMouseTracking(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._drag = None

    def metrics(self):
        sb = self.win.editor.verticalScrollBar()
        total = max(1, sb.maximum() + sb.pageStep())
        return sb, total, max(1, self.height() - 2 * self.PAD)

    def handle(self):
        sb, total, avail = self.metrics()
        hh = max(28.0, sb.pageStep() / total * avail)
        hy = self.PAD + sb.value() / total * avail
        return QRectF(3.5, hy, self.W - 7, min(hh, avail))

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        paint_pane(self, p, r, r.width() / 2, self.win)
        sb, total, avail = self.metrics()
        acc = QColor(self.win.th["accent"])
        hc = QColor(acc)
        hc.setAlpha(110)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(hc)
        hr = self.handle()
        p.drawRoundedRect(hr, hr.width() / 2, hr.width() / 2)
        mark = QColor(self.win.th["text"])
        mark.setAlpha(190)
        p.setBrush(mark)
        for n, y in self.win.editor.page_breaks_y:
            yy = self.PAD + y / total * avail
            p.drawRoundedRect(QRectF(3, yy - 2.5, self.W - 6, 5), 2.5, 2.5)

    def page_at(self, y):
        sb, total, avail = self.metrics()
        doc_y = (y - self.PAD) / avail * total
        return 1 + sum(1 for _, by in self.win.editor.page_breaks_y if by <= doc_y)

    def mousePressEvent(self, e):
        y = e.position().y()
        sb, total, avail = self.metrics()
        hr = self.handle()
        if not hr.contains(QPointF(self.W / 2, y)):
            sb.setValue(int((y - self.PAD) / avail * total - sb.pageStep() / 2))
            hr = self.handle()
        self._drag = (y, sb.value())

    def mouseMoveEvent(self, e):
        y = e.position().y()
        if self._drag:
            sb, total, avail = self.metrics()
            y0, v0 = self._drag
            sb.setValue(int(v0 + (y - y0) / avail * total))
        QToolTip.showText(e.globalPosition().toPoint(), f"стр. {self.page_at(y)}", self)

    def mouseReleaseEvent(self, e):
        self._drag = None

    def wheelEvent(self, e):
        QApplication.sendEvent(self.win.editor.verticalScrollBar(), e)


class BoardViewport(QWidget):
    """Область доски, где живут стикеры. Прокручивается вместе с текстом."""
    def __init__(self, win, parent):
        super().__init__(parent)
        self.win = win
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.setAcceptDrops(True)

    def wheelEvent(self, e):
        QApplication.sendEvent(self.win.editor.verticalScrollBar(), e)

    def mousePressEvent(self, e):
        self.setFocus()
        self.win.stickers.select(None)

    def contextMenuEvent(self, e):
        self.win.show_board_menu(self.mapToGlobal(e.pos()), e.pos())

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls() or e.mimeData().hasImage():
            e.acceptProposedAction()

    def dropEvent(self, e):
        md = e.mimeData()
        pos = e.position().toPoint()
        done = False
        for u in md.urls() if md.hasUrls() else []:
            if u.isLocalFile() and Path(u.toLocalFile()).suffix.lower() in IMG_EXTS:
                self.win.stickers.add_photo(u.toLocalFile(), pos)
                pos += QPoint(20, 20)
                done = True
        if not done and md.hasImage():
            img = QImage(md.imageData())
            if not img.isNull():
                self.win.stickers.add_photo(img, pos)
                done = True
        if done:
            e.acceptProposedAction()

    def paintEvent(self, e):
        if self.win.stickers.items:
            return
        p = QPainter(self)
        f = QFont(self.win.ui_font)
        f.setPointSizeF(9.5)
        p.setFont(f)
        p.setPen(QColor(self.win.th["muted"]))
        p.drawText(QRectF(20, 70, self.width() - 40, 120), Qt.AlignmentFlag.AlignHCenter | Qt.TextFlag.TextWordWrap,
                   "Доска для заметок, фото и ссылок.\nНажми «+» или перетащи сюда картинку.\n\n"
                   "Потяни кружок на краю стикера к строке текста — и они будут связаны линией.")


class Board(QWidget):
    """Левая колонка: непрозрачная доска со стикерами и кнопкой «+», которая всегда на месте."""
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.setObjectName("Board")
        self.vp = BoardViewport(win, self)
        self.plus = QToolButton(self)
        self.plus.setObjectName("AddBtn")
        self.plus.setToolTip("Добавить на доску: заметку, список дел, дедлайн, фото, ссылку, стрелку")
        self.plus.setFixedSize(34, 34)
        self.plus.setIconSize(QSize(18, 18))
        self.plus.setCursor(Qt.CursorShape.PointingHandCursor)
        self.plus.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.plus.clicked.connect(lambda: win.show_board_menu(self.plus.mapToGlobal(self.plus.rect().bottomLeft()
                                                                                    + QPoint(0, 6))))

    def radius(self):
        return 18 * getattr(self.win, "rf", 1.0)

    def resizeEvent(self, e):
        self.vp.setGeometry(self.rect())
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), self.radius(), self.radius())
        from PySide6.QtGui import QRegion
        self.vp.setMask(QRegion(path.toFillPolygon().toPolygon()))
        self.plus.move(10, 10)
        self.plus.raise_()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        th = self.win.th
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        p.setPen(QPen(QColor(th["border"]), 1))
        p.setBrush(QColor(th["page"]))
        p.drawRoundedRect(r, self.radius(), self.radius())


class LinkOverlay(QWidget):
    """Прозрачный слой поверх доски и страницы: рисует все стрелки, поэтому они переходят с доски на текст."""
    def __init__(self, win, parent):
        super().__init__(parent)
        self.win = win
        self.temp = None
        self.geos = {}
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)
        self.hit = ConnectorHit(win, parent)

    def local(self, w, pt):
        """Координаты виджета w → координаты слоя (слой лежит в родителе в точке 0,0)."""
        return w.mapTo(self.parentWidget(), pt) - self.pos()

    def sticker_rect(self, s):
        m = Sticker.MARGIN
        return QRectF(QPointF(self.local(s, QPoint(m, m))), QSizeF(s.width() - 2 * m, s.height() - 2 * m))

    def set_temp(self, a, b=None):
        self.temp = (a, b) if a is not None else None
        self.refresh()

    def setGeometry(self, *a):
        super().setGeometry(*a)
        self.hit.setGeometry(*a)

    def raise_(self):
        super().raise_()
        self.hit.raise_()

    def refresh(self):
        self.update()

    def text_point(self, cur):
        ed = self.win.editor
        vp = ed.viewport()
        cr = ed.cursorRect(cur)
        y = cr.center().y()
        clipped = y < 0 or y > vp.height()
        y = max(6, min(vp.height() - 6, y))
        x = int(ed.document().documentMargin() - 14)
        return QPointF(self.local(vp, QPoint(x, y))), clipped

    @staticmethod
    def curve(a, b):
        path = QPainterPath(a)
        dx = max(40.0, abs(b.x() - a.x()) * 0.5)
        path.cubicTo(QPointF(a.x() + dx, a.y()), QPointF(b.x() - dx, b.y()), b)
        return path

    def paintEvent(self, e):
        win = self.win
        if not hasattr(win, "stickers"):
            return
        L = win.stickers
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        ed, board = win.editor, win.board
        top = max(self.local(board, QPoint(0, 0)).y(), self.local(ed, QPoint(0, 0)).y())
        bottom = min(self.local(board, QPoint(0, board.height())).y(), self.local(ed, QPoint(0, ed.height())).y())
        left = self.local(board, QPoint(0, 0)).x()
        right = self.local(ed, QPoint(ed.width(), 0)).x()
        clip = QRect(left, top, right - left, bottom - top)
        p.setClipRect(clip)
        acc = QColor(win.th["accent"])
        self.geos = {}
        mask_paths, handles = [], []
        sel = L.selected if isinstance(L.selected, Connector) else None
        hover = self.hit.hover
        for c in (L.conns if win.show_links else [x for x in L.conns if x is sel]):
            try:
                g = c.geometry()
            except Exception:
                continue
            br = g["path"].boundingRect().adjusted(-20, -20, 20, 20)
            if not br.intersects(QRectF(clip)):
                continue
            self.geos[c.d["id"]] = g
            w = float(c.d.get("width", 2))
            col = QColor(c.d.get("color") or acc)
            col.setAlphaF(max(0.1, min(1.0, float(c.d.get("opacity", 1.0)))) * (0.6 if g["clipped"] else 1.0))
            if c is sel or c is hover:
                glow = QColor(acc)
                glow.setAlpha(55)
                p.setPen(QPen(glow, w + 8, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
                p.setBrush(Qt.BrushStyle.NoBrush)
                p.drawPath(g["path"])
            style = {"dash": Qt.PenStyle.DashLine, "dot": Qt.PenStyle.DotLine}.get(c.d.get("line"), Qt.PenStyle.SolidLine)
            if g["clipped"]:
                style = Qt.PenStyle.DashLine
            p.setPen(QPen(col, w, style, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawPath(g["path"])
            draw_arrow_head(p, g["p2"], g["t2"], w, col, c.d.get("head_end", "arrow"))
            draw_arrow_head(p, g["p1"], g["t1"], w, col, c.d.get("head_start", "none"))
            mask_paths.append((g["path"], w))
            if c is sel:
                pts = [(g["p1"], False), (g["p2"], False)]
                if c.d.get("shape") != "straight":
                    pts.append((g["handle"], True))
                for q, filled in pts:
                    p.setPen(QPen(acc, 2))
                    p.setBrush(acc if filled else QColor("#FFFFFF"))
                    p.drawEllipse(q, 6.5, 6.5)
                    handles.append(q)
            else:
                handles += [g["p1"], g["p2"]]
        if self.temp and self.temp[1] is not None:
            a, b = self.temp
            p.setPen(QPen(acc, 2, Qt.PenStyle.DashLine, Qt.PenCapStyle.RoundCap))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawPath(self.curve(QPointF(a), QPointF(b)))
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(acc)
            p.drawEllipse(QPointF(b), 4.5, 4.5)
        p.end()
        self.hit.set_region(mask_paths, [q for q in handles if clip.contains(q.toPoint())])


class GlassBar(QFrame):
    """Полупрозрачная панель: под ней размытый фон окна (как «стекло» в iOS)."""
    PAD = 0

    def __init__(self, win, radius=None):
        super().__init__()
        self.win = win
        self.radius = radius
        self.setObjectName("GlassBar")

    def enabled(self):
        return self.win.settings.value("glass", "true") in (True, "true")

    def schedule(self):
        self.update()

    def refresh_blur(self):
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        rf = getattr(self.win, "rf", 1.0)
        radius = (min(r.height() / 2, 24) if self.radius is None else self.radius) * rf
        radius = min(radius, r.height() / 2, r.width() / 2)
        paint_pane(self, p, r, radius, self.win)


class TabStrip(GlassBar):
    """Вкладки открытых документов. Видны, когда открыто больше одного."""
    def __init__(self, win):
        super().__init__(win)
        self.setFixedHeight(38)
        self.h = QHBoxLayout(self)
        self.h.setContentsMargins(5, 4, 5, 4)
        self.h.setSpacing(3)

    def rebuild(self):
        win = self.win
        while self.h.count():
            it = self.h.takeAt(0)
            w = it.widget()
            if w:
                w.hide()
                w.setParent(None)
                w.deleteLater()
        th = win.th
        for i, t in enumerate(win.tabs):
            current = i == win.cur_tab
            path = win.path if current else t.get("path")
            doc = win.doc() if current else t.get("doc")
            name = Path(path).stem if path else "Без названия"
            dirty = doc is not None and doc.isModified()
            box = QWidget()
            box.setObjectName("TabCur" if current else "Tab")
            box.setAttribute(Qt.WidgetAttribute.WA_StyledBackground)
            bl = QHBoxLayout(box)
            bl.setContentsMargins(10, 0, 2, 0)
            bl.setSpacing(0)
            lab = QPushButton()
            lab.setObjectName("TabLabel")
            lab.setCursor(Qt.CursorShape.PointingHandCursor)
            fm = lab.fontMetrics()
            lab.setText(fm.elidedText(name, Qt.TextElideMode.ElideMiddle, 150) + ("  •" if dirty else ""))
            lab.setToolTip(path or "Новый документ — ещё не сохранён")
            lab.clicked.connect(lambda _=False, k=i: win.switch_tab(k))
            bl.addWidget(lab)
            x = QToolButton()
            x.setObjectName("TabClose")
            x.setFixedSize(22, 22)
            x.setIconSize(QSize(12, 12))
            x.setIcon(QIcon(svg_pixmap(SVG["close"], th["muted"], size=12, stroke=2.2)))
            x.setToolTip("Закрыть вкладку  (Ctrl+W)")
            x.setCursor(Qt.CursorShape.PointingHandCursor)
            x.clicked.connect(lambda _=False, k=i: win.close_tab(k))
            bl.addWidget(x)
            box.setFixedHeight(30)
            box.mousePressEvent = (lambda e, k=i: win.close_tab(k) if e.button() == Qt.MouseButton.MiddleButton
                                   else win.switch_tab(k))
            self.h.addWidget(box)
        plus = QToolButton()
        plus.setObjectName("TabClose")
        plus.setFixedSize(28, 28)
        plus.setIcon(QIcon(svg_pixmap(SVG["plus"], th["text"], size=16, stroke=2.2)))
        plus.setToolTip("Новый документ во вкладке  (Ctrl+N)")
        plus.setCursor(Qt.CursorShape.PointingHandCursor)
        plus.clicked.connect(lambda: win.new_document())
        self.h.addWidget(plus)
        self.h.addStretch()
        self.setVisible(len(win.tabs) > 1)

    def refresh(self):
        self.rebuild()


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
HEAD_KINDS = [("Нет", "none"), ("Точка", "dot"), ("Стрелка", "arrow"), ("Открытая", "open"), ("Круг", "circle"),
              ("Ромб", "diamond"), ("Черта", "bar")]


class Inspector(GlassBar):
    """Стеклянная панель справа: все настройки выбранной стрелки или стикера."""
    WIDTH = 336

    def __init__(self, win):
        super().__init__(win, radius=22)
        self.setObjectName("Inspector")
        self.item = None
        self.mode = "doc"
        self.spin = {}
        from PySide6.QtWidgets import QScrollArea
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 8, 8, 10)
        outer.setSpacing(4)
        head = QHBoxLayout()
        self.title = QLabel("")
        self.title.setObjectName("DialogTitle")
        head.addWidget(self.title)
        head.addStretch()
        self.head_btn = RoundButton("", size=30)
        self.head_btn.spec = ("svg", "chev")
        win.icon_buttons.append(self.head_btn)
        self.head_btn.clicked.connect(self.head_clicked)
        head.addWidget(self.head_btn)
        outer.addLayout(head)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll.setStyleSheet("QScrollArea{background:transparent;} QScrollArea > QWidget > QWidget{background:transparent;}")
        outer.addWidget(self.scroll, 1)

    def auto(self):
        return self.win.settings.value("inspector_auto", "true") in (True, "true")

    def expanded(self):
        return self.mode == "item" or self.win.settings.value("doc_panel_open", "false") in (True, "true")

    def head_clicked(self):
        if self.mode == "item":
            self.win.stickers.select(None)
        else:
            self.win.settings.setValue("doc_panel_open", "false" if self.expanded() else "true")
            self.rebuild()

    def mousePressEvent(self, e):
        if self.mode == "doc" and e.position().y() < 48:
            self.head_clicked()

    def hide_panel(self):
        self.item = None
        self.mode = "doc"
        self.rebuild()

    def place(self):
        lay = getattr(self.win, "right_lay", None)
        if lay is not None:
            open_ = self.expanded()
            lay.setStretch(0, 1 if open_ else 0)
            lay.setStretch(1, 0 if open_ else 1)
        self.update()

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
        self.mode = "item"
        self.rebuild()

    def rebuild(self):
        it = self.item
        th = self.win.th
        if it is None:
            self.mode = "doc"
            self.title.setText("Документ")
            open_ = self.expanded()
            self.head_btn.spec = ("svg", "up" if open_ else "down")
            self.head_btn.setToolTip("Свернуть" if open_ else "Настройки документа и вида")
            self.head_btn.setIcon(make_icon(self.head_btn.spec, th))
            self.scroll.setVisible(open_)
            if open_:
                box = QWidget()
                box.setMaximumWidth(self.WIDTH - 30)
                lay = QVBoxLayout(box)
                lay.setContentsMargins(4, 0, 6, 8)
                lay.setSpacing(6)
                self.doc_fields(lay)
                lay.addStretch()
                self.scroll.setWidget(box)
            self.place()
            return
        self.head_btn.spec = ("svg", "close")
        self.head_btn.setToolTip("Снять выделение")
        self.head_btn.setIcon(make_icon(self.head_btn.spec, th))
        self.scroll.setVisible(True)
        pos = self.scroll.verticalScrollBar().value()
        box = QWidget()
        box.setMaximumWidth(self.WIDTH - 30)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(4, 0, 6, 8)
        lay.setSpacing(6)
        self.spin = {}
        if isinstance(it, Connector):
            self.title.setText("Стрелка")
            self.arrow_fields(lay, it)
        else:
            self.title.setText({"note": "Заметка", "photo": "Фото", "link": "Ссылка",
                                "deadline": "Дедлайн"}.get(it.d.get("type"), "Стикер"))
            self.sticker_fields(lay, it)
        lay.addStretch()
        self.scroll.setWidget(box)  # старое содержимое QScrollArea удаляет сам
        self.scroll.verticalScrollBar().setValue(pos)
        QTimer.singleShot(0, self.place)

    def doc_fields(self, lay):
        win, s = self.win, self.win.settings
        text = win.doc().toPlainText()
        words = len(re.findall(r"\w+", text))
        notes = len(win.stickers.items)
        pages = len(win.editor.page_breaks_y) + 1
        info = QLabel(f"{pages} стр.  ·  {words} {plural(words, 'слово', 'слова', 'слов')}  ·  "
                      f"{notes} {plural(notes, 'стикер', 'стикера', 'стикеров')}")
        info.setObjectName("Muted")
        lay.addWidget(info)
        fmt = QLabel(f"Формат: {win.fmt.upper()}" + (f" · {win.encoding.upper()}" if win.fmt == "txt" else ""))
        fmt.setObjectName("Muted")
        lay.addWidget(fmt)
        self._section(lay, "ТЕМА")
        self._seg(lay, [("Светлая", "light"), ("Тёмная", "dark"), ("Авто", "system")], win.theme_mode,
                  win.set_theme_mode)
        self._section(lay, "ФОН ОКНА")
        wm = s.value("wallpaper", "gradient")
        self._seg(lay, [("Нет", "none"), ("Градиент", "gradient"), ("Фото…", "photo")],
                  wm if wm in ("none", "gradient") else "photo", win.set_wallpaper_mode)
        self._section(lay, "ПЛОТНОСТЬ ПАНЕЛЕЙ")
        self._slider(lay, 5, 100, panel_alpha(win) * 100,
                     lambda v: (s.setValue("glass_tint", v / 100), win.wall.update()), "%")
        gl = QCheckBox("Размытие под панелями (стекло)")
        gl.setChecked(s.value("glass", "true") in (True, "true"))
        gl.toggled.connect(lambda on: (s.setValue("glass", "true" if on else "false"), win.wall.update()))
        lay.addWidget(gl)
        self._section(lay, "ЦВЕТ СТРАНИЦЫ И ДОСКИ")
        key = f"page_{win.theme}"
        self._colors(lay, [c for c in PAGE_PRESETS[win.theme] if c], s.value(key, ""),
                     lambda c: (s.setValue(key, c or ""), win.apply_theme()), allow_none=True)
        self._section(lay, "ПОКАЗЫВАТЬ")
        for k, label in (("rulers", "Линейку отступов"), ("page_breaks", "Где заканчиваются страницы"),
                         ("links", "Линии от стикеров к тексту")):
            cb = QCheckBox(label)
            cb.setChecked(s.value(k, "true") in (True, "true"))
            cb.toggled.connect(lambda on, kk=k: (s.setValue(kk, "true" if on else "false"), win.apply_theme()))
            lay.addWidget(cb)
        self._section(lay, "ШИРИНА ДОСКИ")
        self._slider(lay, 180, 520, float(s.value("board_width", 300)),
                     lambda v: (s.setValue("board_width", v), win.apply_layout()), " px")
        lay.addSpacing(4)
        self._buttons(lay, [("Все настройки…", win.open_settings), ("Экспорт…", win.show_export_menu)])

    def arrow_fields(self, lay, a):
        d = a.d
        info = QLabel(f"Откуда: {a.snippet(1)}\nКуда: {a.snippet(2)}")
        info.setObjectName("Muted")
        info.setWordWrap(True)
        lay.addWidget(info)
        self._section(lay, "ЦВЕТ")
        self._colors(lay, ARROW_COLORS, d.get("color"), lambda c: a.set("color", c or self.win.th["accent"]))
        self._section(lay, "ПРОЗРАЧНОСТЬ")
        self._slider(lay, 10, 100, float(d.get("opacity", 1)) * 100, lambda v: a.set("opacity", v / 100), "%")
        self._section(lay, "ТОЛЩИНА")
        self._slider(lay, 1, 12, float(d.get("width", 2)), lambda v: a.set("width", float(v)), " px")
        self._section(lay, "ЛИНИЯ")
        self._seg(lay, [("Сплошная", "solid"), ("Пунктир", "dash"), ("Точки", "dot")],
                  d.get("line", "solid"), lambda v: a.set("line", v))
        self._section(lay, "ФОРМА")
        self._seg(lay, [("Плавная", "curve"), ("Прямая", "straight"), ("Ступенька", "elbow")],
                  d.get("shape", "curve"), lambda v: a.set("shape", v))
        if d.get("shape", "curve") == "curve":
            self._section(lay, "ИЗГИБ")
            self._slider(lay, 0, 250, float(d.get("curv", 1.0)) * 100, lambda v: a.set("curv", v / 100), "%")
        self._section(lay, "КОНЦЫ")
        self._combo(lay, "Начало", HEAD_KINDS, d.get("head_start", "none"), lambda v: a.set("head_start", v))
        self._combo(lay, "Конец", HEAD_KINDS, d.get("head_end", "arrow"), lambda v: a.set("head_end", v))
        hint = QLabel("Тяни белые кружки на концах — к стикеру или строке текста. "
                      "Синий кружок посередине сдвигает изгиб, двойной клик по нему — сброс.")
        hint.setObjectName("Muted")
        hint.setWordWrap(True)
        lay.addWidget(hint)
        lay.addSpacing(4)
        self._buttons(lay, [("Развернуть", a.reverse), ("Сбросить изгиб", a.straighten),
                            ("Отцепить концы", a.detach), ("Удалить", a.delete)])

    def sticker_fields(self, lay, st):
        d = st.d
        typ = d.get("type")
        if typ == "deadline":
            self._section(lay, "ЧТО СДЕЛАТЬ")
            ed = QLineEdit(d.get("title", ""))
            ed.setPlaceholderText("Дедлайн")
            ed.setMinimumHeight(32)
            ed.textChanged.connect(lambda t: (d.__setitem__("title", t), st.update(), st.layer.mark()))
            lay.addWidget(ed)
            when = st.when()
            dl = QLabel("Когда: " + (deadline_text(when) if when else "не указано"))
            dl.setObjectName("Muted")
            lay.addWidget(dl)
            done = QCheckBox("Сделано")
            done.setChecked(bool(d.get("done")))
            done.toggled.connect(lambda on: (d.__setitem__("done", on), st.update(), st.layer.mark()))
            lay.addWidget(done)
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
        self._section(lay, "ВЫРАВНИВАНИЕ НА ДОСКЕ")
        self._buttons(lay, [("⇤ Слева", lambda: self._align(st, "left")), ("По центру", lambda: self._align(st, "center")),
                            ("Справа ⇥", lambda: self._align(st, "right")), ("Наверх экрана", lambda: self._align(st, "top")),
                            ("Вперёд", lambda: st.layer.bring(st, True)),
                            ("Назад", lambda: st.layer.bring(st, False))])
        lay.addSpacing(4)
        extra = [("Стрелка отсюда", st.start_arrow)]
        if typ == "deadline":
            extra.append(("Изменить…", st.edit_deadline))
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
        margin = 12 - Sticker.MARGIN
        if where == "left":
            st.d["x"] = margin
        elif where == "center":
            st.d["x"] = (vw - st.d["w"]) / 2
        elif where == "right":
            st.d["x"] = vw - st.d["w"] - margin
        elif where == "top":
            st.d["y"] = st.layer.scroll() + 52
        st.layer.place(st)
        st.layer.mark()

    def refresh_values(self):
        it = self.item
        if it is None or not self.isVisible() or isinstance(it, Connector):
            return
        for key, sb in self.spin.items():
            sb.blockSignals(True)
            sb.setValue(int(it.d[key]))
            sb.blockSignals(False)


SAVE_FORMATS = [("Word (.docx) — сохраняет оформление", "docx"), ("Простой текст (.txt) — без оформления", "txt"),
                ("Markdown (.md)", "md"), ("Веб-страница (.html)", "html")]


def has_rich_format(qdoc, fmt):
    """Есть ли в тексте оформление, которое потеряется в .txt (или в .md)."""
    md = fmt == "md"
    b = qdoc.begin()
    while b.isValid():
        bf = b.blockFormat()
        if not md and (bf.headingLevel() or b.textList() is not None):
            return True
        al = bf.alignment()
        if al & (Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignJustify):
            return True
        if abs(bf.leftMargin()) > 0.5 or abs(bf.textIndent()) > 0.5:
            return True
        it = b.begin()
        while not it.atEnd():
            f = it.fragment()
            it += 1
            if not f.isValid() or not f.text().strip():
                continue
            cf = f.charFormat()
            if cf.isImageFormat():
                return True
            if not md and (cf.fontWeight() >= 600 or cf.fontItalic() or cf.fontStrikeOut() or cf.isAnchor()):
                return True
            if cf.underlineStyle() != QTextCharFormat.UnderlineStyle.NoUnderline and not cf.isAnchor():
                return True
            if cf.foreground().style() != Qt.BrushStyle.NoBrush or cf.background().style() != Qt.BrushStyle.NoBrush:
                return True
            if cf.fontFamilies() or (cf.fontPointSize() and abs(cf.fontPointSize() - DEFAULT_SIZE) > 0.1
                                     and not bf.headingLevel()):
                return True
        b = b.next()
    return False


class SaveDialog(BaseDialog):
    """Куда сохранить: сразу в нужный проект или папку (или в любое место на компьютере)."""
    def __init__(self, win, name, fmt, start_dir):
        super().__init__(win, "Сохранить документ", "Сохранить")
        from PySide6.QtWidgets import QTreeWidget, QTreeWidgetItem
        self.win = win
        self.setMinimumWidth(460)
        self.name = self.field("Название", name)
        l = QLabel("Формат")
        l.setObjectName("Muted")
        self.lay.addWidget(l)
        self.fmt = QComboBox()
        for label, key in SAVE_FORMATS:
            self.fmt.addItem(label, key)
        self.fmt.setCurrentIndex([k for _, k in SAVE_FORMATS].index(fmt) if fmt in [k for _, k in SAVE_FORMATS] else 0)
        self.fmt.setStyleSheet(f"background:{win.th['bg']};border:1px solid {win.th['border']};")
        self.fmt.setMinimumHeight(34)
        self.lay.addWidget(self.fmt)
        l2 = QLabel("Куда — выбери проект или папку")
        l2.setObjectName("Muted")
        self.lay.addWidget(l2)
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setMinimumHeight(240)
        self.tree.setIconSize(QSize(18, 18))
        self.tree.setStyleSheet(f"QTreeWidget{{background:{win.th['bg']};border:1px solid {win.th['border']};"
                                f"border-radius:12px;padding:6px;outline:none;}} QTreeWidget::item{{padding:4px 2px;border-radius:8px;}}"
                                f"QTreeWidget::item:selected{{background:{win.th['soft']};color:{win.th['text']};}}")
        self.lay.addWidget(self.tree)
        root = win.browser.root
        self.items = {}
        top = QTreeWidgetItem(["Мои проекты"])
        top.setIcon(0, QIcon(cached_svg("home", win.th["accent"], 18)))
        top.setData(0, Qt.ItemDataRole.UserRole, str(root))
        self.tree.addTopLevelItem(top)
        self.items[str(root)] = top
        self._fill(top, root, 0)
        top.setExpanded(True)
        start = str(start_dir) if start_dir else str(root)
        sel = self.items.get(start)
        if sel is None:  # папка вне проектов (например, файл открыт с рабочего стола)
            sel = QTreeWidgetItem([Path(start).name or start])
            sel.setIcon(0, QIcon(cached_svg("folder", win.th["muted"], 18)))
            sel.setData(0, Qt.ItemDataRole.UserRole, start)
            sel.setToolTip(0, start)
            self.tree.addTopLevelItem(sel)
            self.items[start] = sel
        p = sel.parent()
        while p is not None:
            p.setExpanded(True)
            p = p.parent()
        self.tree.setCurrentItem(sel)
        row = QHBoxLayout()
        nf = QPushButton("Новая папка")
        nf.setObjectName("Ghost")
        nf.clicked.connect(self.new_folder)
        other = QPushButton("Другое место на компьютере…")
        other.setObjectName("Ghost")
        other.clicked.connect(self.pick_other)
        row.addWidget(nf)
        row.addWidget(other)
        row.addStretch()
        self.lay.addLayout(row)
        self.finish()
        self.name.setFocus()
        self.name.selectAll()

    def _fill(self, parent_item, folder, depth):
        from PySide6.QtWidgets import QTreeWidgetItem
        if depth > 6:
            return
        try:
            subs = sorted([Path(e.path) for e in os.scandir(folder) if e.is_dir() and visible(e.name)],
                          key=lambda p: p.name.casefold())
        except OSError:
            return
        for d in subs:
            it = QTreeWidgetItem([d.name])
            it.setIcon(0, QIcon(cached_svg("folder", folder_color(d), 18)))
            it.setData(0, Qt.ItemDataRole.UserRole, str(d))
            parent_item.addChild(it)
            self.items[str(d)] = it
            self._fill(it, d, depth + 1)

    def folder(self):
        it = self.tree.currentItem()
        return Path(it.data(0, Qt.ItemDataRole.UserRole)) if it else self.win.browser.root

    def new_folder(self):
        from PySide6.QtWidgets import QTreeWidgetItem
        dlg = NameDialog(self, "Новая папка", "Создать", "Новая папка", with_color=True)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        p = unique_path(self.folder(), dlg.value())
        try:
            p.mkdir(parents=True)
            set_folder_color(p, dlg.color)
        except OSError as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось создать папку:\n{e}")
            return
        parent = self.tree.currentItem()
        it = QTreeWidgetItem([p.name])
        it.setIcon(0, QIcon(cached_svg("folder", dlg.color, 18)))
        it.setData(0, Qt.ItemDataRole.UserRole, str(p))
        parent.addChild(it)
        parent.setExpanded(True)
        self.items[str(p)] = it
        self.tree.setCurrentItem(it)

    def pick_other(self):
        from PySide6.QtWidgets import QTreeWidgetItem
        d = QFileDialog.getExistingDirectory(self, "Папка на компьютере", str(self.folder()))
        if not d:
            return
        it = self.items.get(d)
        if it is None:
            it = QTreeWidgetItem([Path(d).name or d])
            it.setIcon(0, QIcon(cached_svg("folder", self.win.th["muted"], 18)))
            it.setData(0, Qt.ItemDataRole.UserRole, d)
            it.setToolTip(0, d)
            self.tree.addTopLevelItem(it)
            self.items[d] = it
        self.tree.setCurrentItem(it)

    def result(self):
        fmt = self.fmt.currentData()
        return self.folder() / f"{safe_name(self.name.text())}.{fmt}", fmt


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
        self.tabs = []      # открытые документы: doc, path, fmt, encoding, newline, стикеры
        self.cur_tab = -1

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
        self.rf = 1.0
        self.show_links = True
        self.wall = WallRoot(self)
        self.setCentralWidget(self.wall)
        wl = QVBoxLayout(self.wall)
        wl.setContentsMargins(0, 0, 0, 0)
        self.stack = QStackedWidget()
        self.stack.setObjectName("Root")
        wl.addWidget(self.stack)

        self.browser = BrowserPage(self)
        self.stack.addWidget(self.browser)

        self.editor_page = QWidget()
        self.editor_page.setObjectName("Root")
        v = QHBoxLayout(self.editor_page)
        v.setContentsMargins(16, 12, 16, 12)
        v.setSpacing(0)
        left_pill, right_pill = self._build_appbar()
        self._build_fmtbar()
        findbar = self._build_findbar()
        status = self._build_statusbar()
        v.addStretch()
        v.addWidget(self._build_editor_area(left_pill, right_pill, findbar, status), 100)
        v.addStretch()
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
        left = GlassBar(self)
        left.setObjectName("GlassBar")
        h = QHBoxLayout(left)
        h.setContentsMargins(6, 5, 6, 5)
        h.setSpacing(2)
        h.addWidget(self._btn("К проектам  (Alt+←)", ("svg", "back"), self.go_home))
        tcol = QVBoxLayout()
        tcol.setSpacing(0)
        tcol.setContentsMargins(2, 0, 2, 0)
        self.title_label = QLabel("Без названия")
        self.title_label.setObjectName("DocTitle")
        self.sub_label = QLabel("Новый документ")
        self.sub_label.setObjectName("DocSub")
        for l in (self.title_label, self.sub_label):
            l.setMinimumWidth(10)
            l.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        tcol.addWidget(self.title_label)
        tcol.addWidget(self.sub_label)
        h.addLayout(tcol, 1)
        h.addWidget(self._btn("Новый документ  (Ctrl+N)", ("svg", "new"), lambda: self.new_document()))
        h.addWidget(self._btn("Открыть  (Ctrl+O)", ("svg", "open"), self.open_dialog))
        self.left_pill = left

        right = GlassBar(self)
        h = QHBoxLayout(right)
        h.setContentsMargins(6, 5, 6, 5)
        h.setSpacing(2)
        self.undo_btn = self._btn("Отменить  (Ctrl+Z)", ("svg", "undo"), lambda: self.editor.undo())
        self.redo_btn = self._btn("Повторить  (Ctrl+Y)", ("svg", "redo"), lambda: self.editor.redo())
        h.addWidget(self.undo_btn)
        h.addWidget(self.redo_btn)
        h.addWidget(self._btn("Найти и заменить  (Ctrl+F)", ("svg", "find"), self.toggle_find))
        self.theme_btn = self._btn("Тёмная / светлая тема", ("svg", "moon"), self.toggle_theme)
        h.addWidget(self.theme_btn)
        h.addWidget(self._btn("Все настройки: цвета, фон, шрифты", ("svg", "tune"), self.open_settings))

        self.more_btn = self._btn("Ещё", ("svg", "more"))
        self.more_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.more_btn.setObjectName("NoArrow")
        menu = styled_menu(self)
        menu.addAction("Сохранить как…", self.save_as).setShortcut("Ctrl+Shift+S")
        self.export_menu_ref = export_menu(menu, self.export_as, self.th)
        menu.addMenu(self.export_menu_ref)
        menu.addAction("Печать…", self.print_doc).setShortcut("Ctrl+P")
        menu.addSeparator()
        self.view_actions = {}
        vm = styled_menu(menu)
        vm.setTitle("Вид")
        for key, default, label in (("rulers", "true", "Линейка отступов"),
                                    ("page_breaks", "true", "Где заканчиваются страницы"),
                                    ("links", "true", "Линии от стикеров к тексту")):
            act = vm.addAction(label)
            act.setCheckable(True)
            act.toggled.connect(lambda on, k=key: (self.settings.setValue(k, "true" if on else "false"),
                                                   self.apply_theme()))
            self.view_actions[(key, default)] = act
        menu.addMenu(vm)
        menu.addSeparator()
        menu.addAction("Очистить форматирование", self.clear_formatting)
        menu.addAction("Вставить картинку…", self.insert_image)
        menu.addAction("Вставить таблицу…", self.insert_table)
        self.more_btn.setMenu(menu)
        h.addWidget(self.more_btn)
        h.addStretch()
        self.save_btn = QPushButton("Сохранить")
        self.save_btn.setObjectName("Pill")
        self.save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_btn.setToolTip("Ctrl+S")
        self.save_btn.clicked.connect(self.save)
        h.addWidget(self.save_btn)
        self.right_pill = right
        return left, right

    def _sep(self):
        s = QFrame()
        s.setObjectName("Sep")
        s.setFixedSize(1, 22)
        return s

    def _build_fmtbar(self):
        bar = GlassBar(self)
        self.glass = bar
        pad = GlassBar.PAD
        h = FlowLayout(bar, hspacing=2, vspacing=4)
        h.setContentsMargins(10, 5, 10, 5)

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
        m.addAction(make_icon(("svg", "clock"), th), "Дедлайн", self.stickers.add_deadline)
        m.addAction(make_icon(("svg", "image"), th), "Фото-стикер…", self.add_photo_sticker)
        m.addAction(make_icon(("svg", "link"), th), "Ссылка-стикер…", self.stickers.add_link)
        m.addAction(make_icon(("svg", "arrow"), th), "Стрелка", lambda: self.stickers.add_connector())
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

    def _build_editor_area(self, left_pill, right_pill, findbar, status):
        body = QWidget()
        body.setObjectName("EditorArea")
        self.body = self.editor_area = body
        body.installEventFilter(self)
        g = QGridLayout(body)
        g.setContentsMargins(0, 0, 0, 0)
        g.setHorizontalSpacing(12)
        g.setVerticalSpacing(10)
        self.grid = g

        self.editor = Editor()
        self.editor.setObjectName("Editor")
        self.editor.setMinimumWidth(420)
        self.editor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.editor.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
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

        self.hruler = RulerH(self)
        self.board = Board(self)
        self.track = PageTrack(self)
        self.vruler = self.track
        self.right_col = QWidget()
        self.right_lay = QVBoxLayout(self.right_col)
        self.right_lay.setContentsMargins(0, 0, 0, 0)
        self.right_lay.setSpacing(0)

        g.addWidget(left_pill, 0, 0)
        g.addWidget(self.glass, 0, 1)
        g.addWidget(right_pill, 0, 2, 1, 2)
        self.tabbar = TabStrip(self)
        top_mid = QWidget()
        tm = QVBoxLayout(top_mid)
        tm.setContentsMargins(0, 0, 0, 0)
        tm.setSpacing(8)
        tm.addWidget(self.tabbar)
        tm.addWidget(findbar)
        g.addWidget(top_mid, 1, 1)
        g.addWidget(self.hruler, 2, 1)
        g.addWidget(self.board, 3, 0)
        g.addWidget(self.editor, 3, 1)
        g.addWidget(self.track, 3, 2)
        status.setParent(body)
        status.hide()  # счётчик слов теперь в левой верхней панели
        self.status_bar = status
        g.addWidget(self.right_col, 2, 3, 2, 1)
        g.setColumnStretch(1, 1)
        g.setRowStretch(3, 1)

        self.editor.cursorPositionChanged.connect(self.hruler.update)
        self.editor.viewport().installEventFilter(self)
        self.break_timer = QTimer(self)
        self.break_timer.setSingleShot(True)
        self.break_timer.setInterval(500)
        self.break_timer.timeout.connect(self.update_page_breaks)
        self.editor.textChanged.connect(self.break_timer.start)
        self.stickers = StickerLayer(self, self.editor, self.board.vp)
        self.links = LinkOverlay(self, body)
        sb = self.editor.verticalScrollBar()
        sb.valueChanged.connect(lambda _: (self.track.update(), self.links.update()))
        sb.rangeChanged.connect(lambda *_: self.track.update())
        self.editor.textChanged.connect(self.links.update)
        self.inspector = Inspector(self)
        self.right_lay.addWidget(self.inspector)
        self.right_lay.addStretch(1)
        self.count_timer = QTimer(self)
        self.count_timer.setSingleShot(True)
        self.count_timer.setInterval(250)
        self.count_timer.timeout.connect(self.update_counts)
        return body

    def apply_layout(self):
        """Ширины колонок: доска, страница, правая панель."""
        st = self.settings
        bw = int(float(st.value("board_width", 300)))
        pw = int(float(st.value("page_width", 880)))
        rw = Inspector.WIDTH
        self.left_pill.setFixedWidth(bw)
        self.board.setFixedWidth(bw)
        self.right_col.setFixedWidth(rw)
        self.right_pill.setFixedWidth(rw + PageTrack.W + self.grid.horizontalSpacing())
        self.editor.setMaximumWidth(pw)
        self.body.setMaximumWidth(bw + pw + PageTrack.W + rw + 3 * self.grid.horizontalSpacing())
        self.hruler.setVisible(self.show_rulers)
        self.update_title()
        QTimer.singleShot(0, self.place_glass)

    def show_board_menu(self, global_pos, local=None):
        th = self.th
        m = styled_menu(self)
        m.addAction(make_icon(("svg", "sticky"), th), "Заметка", lambda: self.stickers.add_note("note"))
        m.addAction(make_icon(("svg", "checklist"), th), "Список дел", lambda: self.stickers.add_note("todo"))
        m.addAction(make_icon(("svg", "clock"), th), "Дедлайн", self.stickers.add_deadline)
        m.addAction(make_icon(("svg", "image"), th), "Фото…", self.add_photo_sticker)
        m.addAction(make_icon(("svg", "link"), th), "Ссылка…", self.stickers.add_link)
        m.addAction(make_icon(("svg", "arrow"), th), "Стрелка", lambda: self.stickers.add_connector())
        m.exec(global_pos)

    def show_export_menu(self):
        m = export_menu(self, self.export_as, self.th)
        m.exec(QCursor.pos())

    def set_theme_mode(self, k):
        self.theme_mode = k
        self.settings.setValue("theme", k)
        self.theme = self.resolve_theme()
        self.apply_theme()

    def set_wallpaper_mode(self, k):
        if k == "photo":
            path, _ = QFileDialog.getOpenFileName(self, "Фон окна", str(Path.home() / "Pictures"),
                                                  "Изображения (*.png *.jpg *.jpeg *.bmp *.webp)")
            if not path:
                return
            if QImage(path).isNull():
                QMessageBox.warning(self, APP_NAME, "Не получилось открыть эту картинку.")
                return
            dst = fonts_dir().parent / ("wallpaper" + Path(path).suffix.lower())
            for old in fonts_dir().parent.glob("wallpaper.*"):
                try:
                    old.unlink()
                except OSError:
                    pass
            try:
                shutil.copy2(path, dst)
            except OSError as e:
                QMessageBox.warning(self, APP_NAME, f"Не получилось скопировать картинку:\n{e}")
                return
            k = str(dst)
        self.settings.setValue("wallpaper", k)
        self.apply_theme()

    def wallpaper_image(self):
        mode = self.settings.value("wallpaper", "gradient") or "gradient"
        if mode == "none":
            return None, ("none",)
        blur = self.settings.value("glass_blur", 12)
        if mode != "gradient":
            p = Path(mode)
            try:
                key = ("file", str(p), p.stat().st_mtime, self.theme, blur)
            except OSError:
                key = None
            if key is not None:
                if getattr(self, "_wall_file_key", None) != key[:3]:
                    self._wall_file_img = QImage(str(p))
                    self._wall_file_key = key[:3]
                if not self._wall_file_img.isNull():
                    return self._wall_file_img, key
        return gradient_wallpaper(self.th["accent"], self.theme == "dark"), ("gradient", self.th["accent"],
                                                                              self.theme, blur)

    def _build_statusbar(self):
        bar = GlassBar(self)
        bar.setFixedHeight(32)
        h = QHBoxLayout(bar)
        h.setContentsMargins(16, 2, 8, 2)
        self.count_label = QLabel("0 слов")
        self.count_label.setObjectName("Muted")
        self.toast = QLabel("")
        self.toast.setObjectName("Toast")
        self.toast.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.format_chip = QLabel("DOCX")
        self.format_chip.setObjectName("Chip")
        h.addWidget(self.count_label)
        h.addStretch()
        h.addWidget(self.toast)
        h.addStretch()
        right = QWidget()
        right.setMinimumWidth(120)
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
        sc("Ctrl+W", lambda: self.close_tab())
        sc("Ctrl+Tab", lambda: self.next_tab(1))
        sc("Ctrl+Shift+Tab", lambda: self.next_tab(-1))
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
        if obj is getattr(self, "body", None) and ev.type() in (QEvent.Type.Resize, QEvent.Type.Show,
                                                                  QEvent.Type.LayoutRequest):
            self.place_glass()
        elif hasattr(self, "editor") and obj is self.editor.viewport() and ev.type() == QEvent.Type.Resize:
            self.break_timer.start()
            self.hruler.update()
        return super().eventFilter(obj, ev)

    def docked(self):
        return True

    def place_glass(self):
        if hasattr(self, "links"):
            self.links.setGeometry(self.body.rect())
            self.links.raise_()
            self.links.update()

    def set_doc_top(self, top):
        d = self.doc()
        ff = d.rootFrame().frameFormat()
        if abs(ff.topMargin() - top) < 1:
            return
        modified = d.isModified()
        ff.setTopMargin(top)
        d.rootFrame().setFrameFormat(ff)
        d.setModified(modified)
        self.break_timer.start()

    def update_page_breaks(self):
        ed = self.editor
        if False:
            pass
        else:
            try:
                breaks = compute_page_breaks(self.doc())
            except Exception:
                breaks = []
            lay = self.doc().documentLayout()
            out = []
            for n, pos in breaks:
                block = self.doc().findBlock(pos)
                if not block.isValid():
                    continue
                y = lay.blockBoundingRect(block).top()
                bl = block.layout()
                if bl is not None and bl.lineCount():
                    line = bl.lineForTextPosition(pos - block.position())
                    if line.isValid():
                        y += line.y()
                out.append((n, y))
            ed.page_breaks_y = out
        ed.viewport().update()
        self.track.update()
        self.links.update()

    def toggle_setting(self, key, default):
        cur = self.settings.value(key, default) in (True, "true")
        self.settings.setValue(key, "false" if cur else "true")
        self.apply_theme()

    def apply_toolbar_scale(self):
        s = float(self.settings.value("toolbar_scale", 1.0))
        for b in self.glass.findChildren(QToolButton):
            base = b.property("base_size")
            if base is None:
                base = (b.width() if b.minimumWidth() == b.maximumWidth() else 36,
                        b.height() if b.minimumHeight() == b.maximumHeight() else 36)
                b.setProperty("base_size", base)
            b.setFixedSize(round(base[0] * s), round(base[1] * s))
            b.setIconSize(QSize(round(20 * s), round(20 * s)))
        for c in self.glass.findChildren(QComboBox):
            base = c.property("base_w")
            if base is None:
                base = c.width() if c.minimumWidth() == c.maximumWidth() else 120
                c.setProperty("base_w", base)
            c.setFixedWidth(round(base * max(0.8, s)))
            c.setFixedHeight(round(32 * s))
        for sep in self.glass.findChildren(QFrame, "Sep"):
            sep.setFixedSize(1, round(22 * s))

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
        self.board.plus.setIcon(QIcon(svg_pixmap(SVG["plus"], th["accent_text"], size=18, stroke=2.4)))
        if hasattr(self, "stickers"):
            self.stickers.refresh_styles()
        if hasattr(self, "glass"):
            self.glass.schedule()
            QTimer.singleShot(0, self.place_glass)
        if hasattr(self, "export_menu_ref"):
            self.export_menu_ref.setIcon(make_icon(("svg", "export"), th))
        if hasattr(self, "view_actions"):
            for key, act in self.view_actions.items():
                act.blockSignals(True)
                act.setChecked(self.settings.value(key[0], key[1]) in (True, "true"))
                act.blockSignals(False)

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
        self.show_rulers = st.value("rulers", "true") in (True, "true")
        self.show_breaks = st.value("page_breaks", "true") in (True, "true")
        self.show_links = st.value("links", "true") in (True, "true")
        tb_s = float(st.value("toolbar_scale", 1.0))
        tb_font = float(st.value("toolbar_font", 10))
        tb_r = round(18 * tb_s)
        self.rf = float(st.value("ui_round", 1.0))
        ui_size = float(st.value("ui_font", 10))
        self.apply_layout()
        img, key = self.wallpaper_image()
        self.wall.set_source(img, key)
        wall_on = img is not None
        ov = QColor(th["surface"])
        empty_bg = f"rgba({ov.red()},{ov.green()},{ov.blue()},0.82)"
        sheet = f"""
        * {{ font-family: "Segoe UI Variable Text", "Segoe UI", "Inter", sans-serif; font-size: {ui_size:g}pt; color: {th['text']}; }}
        #Root, #AppBar, #FmtWrap, #EditorArea, #StatusBar, #Browser {{ background: transparent; }}
        #EmptyPill {{ background: {empty_bg}; border-radius: 22px; padding: 26px 34px; color: {th['muted']};
                      font-size: 11pt; }}
        #DocTitle {{ font-size: 12pt; font-weight: 600; }}
        #DocSub, #Muted {{ color: {th['muted']}; font-size: 9pt; }}
        #DocSub[toast="true"] {{ color: {th['accent']}; font-weight: 600; }}
        #Toast {{ color: {th['accent']}; font-size: 9pt; font-weight: 600; }}
        #Chip {{ background: {th['surface']}; border: 1px solid {th['border']}; border-radius: 11px;
                 padding: 2px 10px; color: {th['muted']}; font-size: 8.5pt; font-weight: 600; }}
        #FmtBar {{ background: {th['surface']}; border: 1px solid {th['border']}; border-radius: 23px; }}
        #Board QToolButton#AddBtn {{ border-radius: 17px; }}
        QWidget#Tab, QWidget#TabCur {{ border-radius: 15px; }}
        QWidget#TabCur {{ background: {th['page']}; border: 1px solid {th['border']}; }}
        QWidget#Tab:hover {{ background: {th['hover']}; }}
        QPushButton#TabLabel {{ background: transparent; border: none; padding: 4px 4px; color: {th['muted']}; }}
        QWidget#TabCur QPushButton#TabLabel {{ color: {th['text']}; font-weight: 600; }}
        QToolButton#TabClose {{ border-radius: 11px; }}
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
        #Editor {{ font-family: "{DEFAULT_FONT}"; font-size: {DEFAULT_SIZE}pt; background: {th['page']};
                   border: 1px solid {th['border']}; border-radius: 18px; color: {th['doc_text']};
                   selection-background-color: {th['sel']}; selection-color: {th['doc_text']}; }}
        #GlassBar, #Inspector {{ background: transparent; border: none; }}
        #GlassBar QToolButton {{ border-radius: {tb_r}px; }}
        #GlassBar QComboBox, #GlassBar QFontComboBox {{ font-size: {tb_font:g}pt; }}
        QToolButton#AddBtn {{ border-radius: {tb_r}px; }}
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
        self.apply_toolbar_scale()
        self.hruler.update()
        self.track.update()
        self.board.update()
        self.links.update()
        self.wall.update()
        if hasattr(self, "inspector") and self.inspector.mode == "doc":
            self.inspector.rebuild()
        QTimer.singleShot(0, self.place_glass)
        self.break_timer.start()

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
        d.modificationChanged.connect(self.update_title)
        d.setUndoRedoEnabled(False)
        ff = d.rootFrame().frameFormat()
        ff.setTopMargin(40)
        ff.setBottomMargin(360)  # можно прокрутить ниже текста — место для стикеров на доске
        d.rootFrame().setFrameFormat(ff)
        self.apply_spacing_all()
        d.setUndoRedoEnabled(True)
        d.setModified(False)
        self.undo_btn.setEnabled(False)
        self.redo_btn.setEnabled(False)
        self.editor.page_breaks_y = []
        self.break_timer.start()
        self.hruler.update()
        self.inspector.rebuild()

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
        if self.path:
            parent = Path(self.path).parent
            try:
                rel = parent.relative_to(self.browser.root)
                sub = " › ".join(("Проекты",) + rel.parts)
            except ValueError:
                sub = str(parent)
        else:
            sub = "Новый документ"
        # под названием — сколько слов; на пару секунд там же появляются сообщения («Сохранено ✓»)
        toast = getattr(self, "_toast_text", "")
        line2 = toast or getattr(self, "count_text", "")
        self.sub_label.setProperty("toast", bool(toast))
        self._repolish(self.sub_label)
        for lab, text in ((self.title_label, name + ("  •" if dirty else "")), (self.sub_label, line2)):
            w = max(40, lab.width() - 2)
            lab.setText(lab.fontMetrics().elidedText(text, Qt.TextElideMode.ElideMiddle, w))
        self.title_label.setToolTip(f"{name}\n{sub}")
        self.sub_label.setToolTip(sub)
        self.setWindowTitle(f"{'• ' if dirty else ''}{name} — {APP_NAME}")
        if hasattr(self, "tabbar") and self.tabbar.isVisible():
            self.tabbar.rebuild()
        self.format_chip.setText(self.fmt.upper() + (f" · {self.encoding.upper()}" if self.fmt == "txt" else ""))

    def new_document(self, ask=True):
        if self.tabs and self.is_blank_tab():
            self.stack.setCurrentWidget(self.editor_page)
            self.editor.setFocus()
            return
        self.new_tab_slot()
        self.editor.setDocument(QTextDocument(self))
        self.path, self.fmt, self.encoding, self.newline = None, "docx", "utf-8", "\r\n"
        self._target = None
        self.stickers.load(None)
        self.setup_document()
        self._sync_tab()
        self.update_title()
        self.update_counts()
        self.sync_toolbar()
        self.tabbar.rebuild()
        self.editor.setFocus()

    # ── Вкладки ──
    def is_blank_tab(self):
        return (not self.path and not self.doc().isModified() and not self.doc().toPlainText().strip()
                and not self.stickers.items)

    def new_tab_slot(self):
        """Освобождает место под новый документ: текущая вкладка сохраняется в списке."""
        if self.cur_tab >= 0 and self.tabs:
            self._stash()
        self.tabs.append({})
        self.cur_tab = len(self.tabs) - 1

    def _sync_tab(self):
        if 0 <= self.cur_tab < len(self.tabs):
            self.tabs[self.cur_tab].update(doc=self.editor.document(), path=self.path, fmt=self.fmt,
                                           encoding=self.encoding, newline=self.newline)

    def _stash(self):
        if not (0 <= self.cur_tab < len(self.tabs)):
            return
        self.stickers.save()
        self._sync_tab()
        t = self.tabs[self.cur_tab]
        t["stickers"] = self.stickers.dump()
        t["cursor"] = self.editor.textCursor().position()
        t["scroll"] = self.editor.verticalScrollBar().value()

    def tab_index(self, path):
        try:
            rp = Path(path).resolve()
        except OSError:
            return -1
        for i, t in enumerate(self.tabs):
            p = self.path if i == self.cur_tab else t.get("path")
            if p and Path(p).resolve() == rp:
                return i
        return -1

    def switch_tab(self, i):
        if not (0 <= i < len(self.tabs)):
            return
        if i != self.cur_tab:
            self._stash()
            self._activate(i)
        self.stack.setCurrentWidget(self.editor_page)
        self.editor.setFocus()

    def _activate(self, i):
        self.cur_tab = i
        t = self.tabs[i]
        self.editor.setDocument(t["doc"])
        self.path, self.fmt = t.get("path"), t.get("fmt", "docx")
        self.encoding, self.newline = t.get("encoding", "utf-8"), t.get("newline", "\r\n")
        self._target = None
        self.stickers.load_data(t.get("stickers") or [], self.path)
        c = self.editor.textCursor()
        c.setPosition(min(t.get("cursor", 0), self.doc().characterCount() - 1))
        self.editor.setTextCursor(c)
        QTimer.singleShot(0, lambda v=t.get("scroll", 0): self.editor.verticalScrollBar().setValue(v))
        self.undo_btn.setEnabled(self.doc().isUndoAvailable())
        self.redo_btn.setEnabled(self.doc().isRedoAvailable())
        self.editor.page_breaks_y = []
        self.break_timer.start()
        self.update_title()
        self.update_counts()
        self.sync_toolbar()
        self.hruler.update()
        self.inspector.rebuild()
        self.tabbar.rebuild()

    def close_tab(self, i=None):
        i = self.cur_tab if i is None else i
        if not (0 <= i < len(self.tabs)):
            return
        if i != self.cur_tab:
            self.switch_tab(i)
        if not self.maybe_save():
            return
        self.stickers.save()
        self.tabs.pop(i)
        self.cur_tab = -1
        if not self.tabs:
            self.new_document()
            self.go_home()
            return
        self._activate(min(i, len(self.tabs) - 1))

    def next_tab(self, step=1):
        if len(self.tabs) > 1:
            self.switch_tab((self.cur_tab + step) % len(self.tabs))

    def open_dialog(self):
        start = self.settings.value("last_dir", str(Path.home() / "Documents"))
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть", start,
            "Все документы (*.docx *.txt *.md *.markdown *.html *.htm *.log *.csv *.json *.xml *.ini);;"
            "Word (*.docx);;Текст (*.txt);;Markdown (*.md *.markdown);;Веб-страница (*.html *.htm);;Все файлы (*)")
        if path and self.open_path(path):
            self.stack.setCurrentWidget(self.editor_page)

    def open_path_checked(self, path):
        if self.open_path(path):
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
        existing = self.tab_index(p)
        if existing >= 0:  # уже открыт — просто переключаемся
            self.switch_tab(existing)
            return True
        newdoc = QTextDocument(self)
        enc, nl = "utf-8", "\r\n"
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
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Не получилось открыть файл:\n{p.name}\n\n{e}")
            return False
        if not (self.tabs and self.is_blank_tab()):
            self.new_tab_slot()
        self.editor.setDocument(newdoc)
        self.path, self.fmt = str(p), fmt
        self.encoding, self.newline = enc, nl
        self._target = None
        self.stickers.load(str(p))
        self.setup_document()
        self._sync_tab()
        self.tabbar.rebuild()
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
        if self.open_path(path):
            self.stack.setCurrentWidget(self.editor_page)
            self.editor.setFocus()

    def go_home(self):
        # несохранённые вкладки остаются открытыми — к ним можно вернуться
        if self.findbar.isVisible():
            self.toggle_find(False)
        self.stickers.save()
        self._sync_tab()
        target = self.browser.cwd
        if self.path:
            parent = Path(self.path).parent
            ext = getattr(self.browser, "ext", None)
            if self.browser.inside(parent, self.browser.root) or (ext and self.browser.inside(parent, ext[0])):
                target = parent
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
        old, new = Path(old), Path(new)

        def moved(p):
            if not p:
                return p
            try:
                rel = Path(p).relative_to(old)
            except ValueError:
                return p
            return str(new / rel) if str(rel) != "." else str(new)
        for i, t in enumerate(self.tabs):
            if i != self.cur_tab:
                t["path"] = moved(t.get("path"))
        if self.path:
            self.path = moved(self.path)
            self.stickers.path = self.path
        self._sync_tab()
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
        """Перед обновлением: сохранить стикеры, спросить про несохранённые документы."""
        if not self.ask_save_all():
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
        if self.fmt in ("txt", "md") and has_rich_format(self.doc(), self.fmt):
            choice = self.ask_keep_format()
            if choice is None:
                return False
            if choice == "docx":
                return self.convert_to_docx()
        return self.write_file(self.path, self.fmt)

    def ask_keep_format(self):
        box = QMessageBox(self)
        box.setWindowTitle(APP_NAME)
        box.setIcon(QMessageBox.Icon.NoIcon)
        kind = "простом тексте (.txt)" if self.fmt == "txt" else "Markdown (.md)"
        box.setText(f"В {kind} оформление не сохранится")
        box.setInformativeText("Жирный, цвета, шрифты, размеры и выравнивание пропадут, когда файл откроется снова.\n\n"
                               "Сохранить как Word (.docx)? Получится новый файл рядом с этим, со всем оформлением — "
                               "дальше Folio будет работать с ним.")
        word = box.addButton("Сохранить как Word", QMessageBox.ButtonRole.AcceptRole)
        plain = box.addButton("Без оформления", QMessageBox.ButtonRole.DestructiveRole)
        box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(word)
        box.exec()
        c = box.clickedButton()
        return "docx" if c is word else "plain" if c is plain else None

    def convert_to_docx(self):
        old = Path(self.path)
        new = old.with_suffix(".docx")
        if new.exists():
            new = unique_path(old.parent, old.stem, ".docx")
        if not self.write_file(str(new), "docx"):
            return False
        sc = sidecar_path(old)
        if sc.exists():  # стикеры переезжают к новому файлу
            try:
                sc.unlink()
            except OSError:
                pass
        self.path, self.fmt, self.encoding = str(new), "docx", "utf-8"
        self._sync_tab()
        self.update_title()
        self.show_toast("Сохранено как Word ✓")
        return True

    def save_as(self):
        if self.path:
            start_dir = Path(self.path).parent
        else:
            start_dir = self.browser.cwd
            if self.browser.at_root():
                last = self.settings.value("last_save_dir", "")
                if last and Path(last).is_dir():
                    start_dir = Path(last)
        base = Path(self.path).stem if self.path else self.suggest_name()
        fmt = self.fmt
        if fmt in ("txt", "md") and has_rich_format(self.doc(), fmt):
            fmt = "docx"
        dlg = SaveDialog(self, base, fmt, start_dir)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return False
        path, fmt = dlg.result()
        if path.exists() and (not self.path or path.resolve() != Path(self.path).resolve()):
            box = QMessageBox(self)
            box.setWindowTitle(APP_NAME)
            box.setIcon(QMessageBox.Icon.NoIcon)
            box.setText(f"«{path.name}» там уже есть. Заменить?")
            yes = box.addButton("Заменить", QMessageBox.ButtonRole.DestructiveRole)
            box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() is not yes:
                return False
        if not self.write_file(str(path), fmt):
            return False
        self.path, self.fmt = str(path), fmt
        self.settings.setValue("last_dir", str(path.parent))
        if self.browser.inside(path, self.browser.root):
            self.settings.setValue("last_save_dir", str(path.parent))
        self._sync_tab()
        self.update_title()
        self.browser.refresh()
        return True

    def suggest_name(self):
        first = self.doc().begin().text().strip()
        return safe_name(first[:60]) if first else "Без названия"

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
        self.tabbar.rebuild()
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
        self._toast_text = text
        self.update_title()
        QTimer.singleShot(2300, self._clear_toast)

    def _clear_toast(self):
        if not self.toast_timer.isActive():
            self._toast_text = ""
            self.update_title()

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
        self.count_text = f"{words} {plural(words, 'слово', 'слова', 'слов')} · {chars} {plural(chars, 'символ', 'символа', 'символов')}"
        self.update_title()

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
    def ask_save_all(self):
        """Спросить про каждый несохранённый документ во вкладках."""
        for i in range(len(self.tabs)):
            doc = self.doc() if i == self.cur_tab else self.tabs[i].get("doc")
            if doc is not None and doc.isModified():
                self.switch_tab(i)
                if not self.maybe_save():
                    return False
        return True

    def closeEvent(self, e):
        if self.ask_save_all():
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
