#!/usr/bin/env python3
"""Apple-style (SF Symbols-like) line icons for RT-QuIC Editor.
Renders inline SVG strings to QIcon via QSvgRenderer.
"""
from PyQt5.QtCore import Qt, QByteArray, QRectF
from PyQt5.QtGui import QIcon, QPixmap, QPainter
from PyQt5.QtSvg import QSvgRenderer

# Each icon is a 24x24 viewBox, stroke-based (fill none), Apple-style thin line.

SVGS = {
    'folder': "<path d='M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z'/>",
    'export': "<path d='M9 3h6l4 4v12a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z'/><path d='M9 3v4h6'/><path d='M12 16V9'/><path d='M9 12l3-3 3 3'/>",
    'braces': "<path d='M8 4C6 4 5 5 5 7v2c0 1.6-1 2-2 3 1 1 2 1.4 2 3v2c0 2 1 3 3 3M16 4c2 0 3 1 3 3v2c0 1.6 1 2 2 3-1 1-2 1.4-2 3v2c0 2-1 3-3 3'/>",
    'rows': "<path d='M4 6h16M4 12h16M4 18h9'/><path d='M19 15.5v5M16.5 18h5'/>",
    'cols': "<path d='M6 4v16M12 4v16M18 4v9'/><path d='M15.5 19h5M18 16.5v5'/>",
    'trash': "<path d='M4 7h16M9 7V5h6v2M6 7l1 13h10l1-13M10 11v5M14 11v5'/>",
    'clock': "<path d='M12 21a9 9 0 1 1 9-9'/><path d='M12 7v5l3 2'/>",
    'chartbar': "<path d='M4 20h16M7 20v-6M12 20V9M17 20v-9'/>",
    'note': "<path d='M7 3h7l4 4v14H7z'/><path d='M14 3v4h4M9 12h7M9 16h7'/>",
    'doc': "<path d='M7 3h7l4 4v14H7z'/><path d='M14 3v4h4'/>",
    'gear': "<circle cx='12' cy='12' r='3'/><path d='M12 2.5v3M12 18.5v3M2.5 12h3M18.5 12h3M5 5l2.1 2.1M16.9 16.9L19 19M19 5l-2.1 2.1M7.1 16.9L5 19'/>",
    'sparkles': "<path d='M12 3.5l1.9 4.6L18.5 10l-4.6 1.9L12 16.5l-1.9-4.6L5.5 10l4.6-1.9z'/><path d='M19 14.5l.9 2.1 2.1.9-2.1.9-.9 2.1-.9-2.1-2.1-.9 2.1-.9z'/>",
    'sidebar': "<rect x='3' y='4' width='18' height='16' rx='2'/><path d='M9 4v16'/>",
    'table': "<rect x='3' y='4' width='18' height='16' rx='2'/><path d='M3 9h18M3 14h18M12 9v11'/>",
    'eye': "<path d='M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z'/><circle cx='12' cy='12' r='3'/>",
    'plus': "<circle cx='12' cy='12' r='9'/><path d='M12 8v8M8 12h8'/>",
    'doc_plus': "<path d='M8 3h7l4 4v14H8z'/><path d='M15 3v4h4'/><circle cx='16' cy='18' r='3'/><path d='M16 16.5v3M14.5 18h3'/>",
    'chartline': "<path d='M3 17l5-6 4 3 6-8'/><path d='M3 21h18'/>",
    'plate': "<rect x='3' y='4' width='18' height='16' rx='2'/><path d='M3 10h18M3 14h18M9 4v16M15 4v16'/>",
    'sliders': "<path d='M4 8h9M17 8h3M4 16h3M11 16h9'/><circle cx='15' cy='8' r='2'/><circle cx='9' cy='16' r='2'/>",
    'text': "<path d='M6 5h12M12 5v14M9 19h6'/>",
    'axis': "<path d='M20 4v16M20 4l-3 3M20 4l3 3M4 20h16M4 20l3-3M4 20l-3-3'/>",
    'range': "<path d='M3 12h18M3 12l4-4M3 12l4 4M21 12l-4-4M21 12l-4 4'/>",
    'globe': "<circle cx='12' cy='12' r='9'/><path d='M3 12h18M12 3a15 15 0 0 1 0 18M12 3a15 15 0 0 0 0 18'/>",
    'image': "<rect x='3' y='5' width='18' height='14' rx='2'/><circle cx='9' cy='10' r='1.5'/><path d='M3 17l5-4 4 3 4-4 5 5'/>",
    'refresh': "<path d='M20 12a8 8 0 1 1-2.3-5.6M20 4v4h-4'/>",
    'check': "<circle cx='12' cy='12' r='9'/><path d='M8 12l3 3 5-6'/>",
    'xmark': "<circle cx='12' cy='12' r='9'/><path d='M9 9l6 6M15 9l-6 6'/>",
    'undo': "<path d='M9 14L4 9l5-5'/><path d='M4 9h10a6 6 0 0 1 6 6v1'/>",
    'redo': "<path d='M15 14l5-5-5-5'/><path d='M20 9H10a6 6 0 0 0-6 6v1'/>",
    'play': "<path d='M8 5.5v13l10-6.5z'/>",
    'question': "<circle cx='12' cy='12' r='9'/><path d='M9.5 9.5a2.5 2.5 0 1 1 3.4 2.3c-.9.4-1.4 1-1.4 2.2M12 17h.01'/>",
}

_CACHE = {}


def make_icon(name, color='#1a1d23', size=18):
    """Render an icon to a QIcon at the requested pixel size."""
    svg = SVGS.get(name)
    if not svg:
        return QIcon()
    key = (name, color, size)
    if key in _CACHE:
        return _CACHE[key]
    g = ("stroke='%s' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round' fill='none'") % color
    full = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' width='%d' height='%d'>"
            "<g %s>%s</g></svg>") % (size, size, g, svg)
    renderer = QSvgRenderer(QByteArray(full.encode('utf-8')))
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    if not renderer.isValid():
        _CACHE[key] = QIcon()
        return _CACHE[key]
    p = QPainter(pm)
    renderer.render(p, QRectF(0, 0, size, size))
    p.end()
    icon = QIcon(pm)
    _CACHE[key] = icon
    return icon


def svg_html(name, color='currentColor', size=16, weight=1.7):
    """Return an inline <svg> snippet for use inside HTML (viewer)."""
    svg = SVGS.get(name, '')
    return ("<svg viewBox='0 0 24 24' width='%d' height='%d' style='vertical-align:-3px;display:inline-block;'>"
            "<g stroke='%s' stroke-width='%s' stroke-linecap='round' stroke-linejoin='round' fill='none'>%s</g></svg>"
            % (size, size, color, weight, svg))


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtCore import Qt, QByteArray, QRectF
    from PyQt5.QtGui import QPixmap, QPainter
    from PyQt5.QtSvg import QSvgRenderer
    app = QApplication(sys.argv)
    bad = []
    for name in SVGS:
        svg = SVGS[name]
        g = "stroke='#111' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round' fill='none'"
        full = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="18" height="18">'
                '<g ' + g + '>' + svg + '</g></svg>')
        r = QSvgRenderer(QByteArray(full.encode('utf-8')))
        pm = QPixmap(18, 18); pm.fill(Qt.transparent)
        if not r.isValid():
            bad.append(name + '(invalid)'); continue
        p = QPainter(pm); r.render(p, QRectF(0, 0, 18, 18)); p.end()
        img = pm.toImage()
        n = sum(1 for x in range(18) for y in range(18) if img.pixelColor(x, y).alpha() > 0)
        if n == 0:
            bad.append(name + '(blank)')
    print('icons:', len(SVGS), 'bad:', bad if bad else 'none')
