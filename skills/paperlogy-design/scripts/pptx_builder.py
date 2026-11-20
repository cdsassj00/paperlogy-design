# -*- coding: utf-8 -*-
"""
pptx_builder.py — Paperlogy 디자인 언어로 '편집 가능한' PPTX를 만드는 라이브러리.

    import sys; sys.path.insert(0, r"~/.claude/skills/paperlogy-design/scripts")
    from pptx_builder import Deck
    d = Deck(theme="pink")                       # pink | blue | mint | red | yellow
    d.cover("PROPOSAL 2026", "**AI**로 **일하는 방식**을\\n**바꾸다**.", "부제", "발표자 · 2026.10")
    d.save("out/deck.pptx")                      # + out/deck.layouts.json (검증용)
    → python render_pptx.py out/deck.pptx out/png --check

좌표·크기는 HTML 템플릿과 같은 1280×720 px 캔버스 기준(1px = 9525 EMU, 글자 px×0.75 = pt).
텍스트 마크업:  **굵게(키워드)**  ==액센트색==  ~~회색으로 죽이기~~  , 줄바꿈은 \\n
               제목 끝의 '.'은 자동으로 마침표 컬러(--dot).
폰트: 제목·디스플레이 = Paperlogy(굵기별 패밀리), 본문·표 = Pretendard, 영문 라벨 = Paperlogy 9 Black.
배경·구체·후광은 PIL로 PNG를 생성해 깐다(테마 색 반영). 텍스트·도형·차트·표는 전부 네이티브.
"""
import os, re, json, math, copy, tempfile
import numpy as np
from PIL import Image, ImageDraw, ImageFilter
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

PX = 9525
W, H = 1280, 720

THEMES = {
    "pink":   dict(accent="FF2F92", on_dark="FF4FA3", accent2="4947FF", dot="FF4FA3", g1="8A39C0", g2="4A1670", g3="5B7BFF"),
    "blue":   dict(accent="4947FF", on_dark="8E8CFF", accent2="1F9873", dot="8E8CFF", g1="3346D6", g2="121C66", g3="36C3FF"),
    "mint":   dict(accent="239E78", on_dark="53E0B4", accent2="4947FF", dot="53E0B4", g1="1F8F72", g2="0B3A30", g3="3AA0FF"),
    "red":    dict(accent="E53935", on_dark="FF6B66", accent2="5A5B60", dot="FF6B66", g1="B0283A", g2="4A0D18", g3="FF8A3D"),
    "yellow": dict(accent="A07E00", on_dark="FFD84D", accent2="4947FF", dot="FFD84D", g1="8A6A00", g2="2E2400", g3="FFB347"),
}

F_TITLE = {"r": "Paperlogy 4 Regular", "b": "Paperlogy 8 ExtraBold", "l": "Paperlogy 3 Light",
           "t": "Paperlogy 1 Thin", "k": "Paperlogy 9 Black", "m": "Paperlogy 6 SemiBold"}
F_BODY = {"r": "Pretendard", "b": "Pretendard", "m": "Pretendard Medium", "l": "Pretendard Light"}
F_EN = "Paperlogy 9 Black"

WHITE, BODY, SUB, DIM, INK, INK_BODY, LINE, PAPER2 = "FFFFFF", "BEBEBE", "EBEBEB", "3A3B40", "1D1D1F", "595959", "D0D1DD", "EFF0F2"


def E(px):
    return Emu(int(round(px * PX)))


def _hex(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _mix(a, b, t):
    return tuple(int(a[i] * (1 - t) + b[i] * t) for i in range(3))


# ─────────────────────────── 이미지 에셋 (PIL) ───────────────────────────
class Assets:
    def __init__(self, folder, th):
        self.dir, self.th = folder, th
        os.makedirs(folder, exist_ok=True)

    def _path(self, name):
        return os.path.join(self.dir, name + ".png")

    @staticmethod
    def _radial(w, h, cx, cy, rx, ry, stops):
        """stops: [(pos 0..1, (r,g,b), alpha 0..1)] → float RGBA array"""
        y, x = np.mgrid[0:h, 0:w].astype(np.float32)
        d = np.sqrt(((x - cx) / rx) ** 2 + ((y - cy) / ry) ** 2)
        out = np.zeros((h, w, 4), np.float32)
        pos = [s[0] for s in stops]
        for c in range(3):
            out[..., c] = np.interp(d, pos, [s[1][c] for s in stops])
        out[..., 3] = np.interp(d, pos, [s[2] for s in stops])
        return out

    @staticmethod
    def _over(base, layer):
        a = layer[..., 3:4]
        base[..., :3] = base[..., :3] * (1 - a) + layer[..., :3] * a
        return base

    def _save(self, arr, name, alpha=False):
        arr = arr.copy()
        if alpha:
            arr[..., 3] *= 255          # 알파는 0..1로 계산 → 0..255
        arr = np.clip(arr, 0, 255)
        if alpha:
            img = Image.fromarray(arr.astype(np.uint8), "RGBA")
        else:
            img = Image.fromarray(arr[..., :3].astype(np.uint8), "RGB")
        p = self._path(name)
        img.save(p)
        return p

    def bg(self, kind):
        p = self._path("bg_" + kind)
        if os.path.exists(p):
            return p
        w, h = 1920, 1080
        g1, g2, g3 = _hex(self.th["g1"]), _hex(self.th["g2"]), _hex(self.th["g3"])
        base = np.zeros((h, w, 4), np.float32)
        if kind == "aurora_top":
            base = self._over(base, self._radial(w, h, w * .5, -h * .12, w * .62, h * .70,
                              [(0, g1, 1), (.35, g2, 1), (.7, _mix(g2, (0, 0, 0), .7), 1), (1, (0, 0, 0), 1), (9, (0, 0, 0), 1)]))
        elif kind == "aurora_bottom":
            base = self._over(base, self._radial(w, h, w * .62, h * 1.18, w * .70, h * .72,
                              [(0, g3, 1), (.33, g1, 1), (.66, _mix(g2, (0, 0, 0), .4), 1), (1, (0, 0, 0), 1), (9, (0, 0, 0), 1)]))
        elif kind == "nebula":
            base = self._over(base, self._radial(w, h, w * .52, h * .18, w * .32, h * .40, [(0, (56, 200, 150), .40), (1, (0, 0, 0), 0), (9, (0, 0, 0), 0)]))
            base = self._over(base, self._radial(w, h, w * .85, h * .10, w * .40, h * .45, [(0, (40, 80, 200), .45), (1, (0, 0, 0), 0), (9, (0, 0, 0), 0)]))
            base = self._over(base, self._radial(w, h, w * .40, h * .30, w * .30, h * .30, [(0, (120, 255, 220), .15), (1, (0, 0, 0), 0), (9, (0, 0, 0), 0)]))
            # 좌하단 행성 림(오렌지)
            cx, cy, r = w * .10, h * 1.10, h * .42
            y, x = np.mgrid[0:h, 0:w].astype(np.float32)
            d = np.sqrt((x - cx) ** 2 + (y - cy) ** 2)
            ring = np.clip(1 - np.abs(d - r) / 26, 0, 1) * np.clip((cx + r * .2 - x) / (r * .9) + (cy - y) / (r * 1.2), 0, 1)
            lay = np.zeros((h, w, 4), np.float32)
            lay[..., 0], lay[..., 1], lay[..., 2], lay[..., 3] = 255, 140, 50, ring * .85
            base = self._over(base, lay)
        elif kind == "dark_report":
            t = np.linspace(0, 1, h)[:, None]
            v = np.interp(t, [0, .6, 1], [34, 11, 3])
            base[..., 0] = base[..., 1] = base[..., 2] = v
        elif kind in ("nightsky", "night"):
            t = np.linspace(0, 1, h)
            if kind == "nightsky":
                pts = [0, .38, .58, .72, .76, .80, 1]
                cols = [(13, 26, 92), (40, 58, 158), (91, 79, 184), (199, 106, 158), (231, 136, 154), (42, 47, 122), (27, 29, 85)]
            else:
                pts = [0, .55, .70, 1]
                cols = [(5, 7, 15), (11, 17, 36), (18, 26, 51), (4, 6, 12)]
            for c in range(3):
                base[..., c] = np.interp(t, pts, [col[c] for col in cols])[:, None]
            rng = np.random.default_rng(7)
            for _ in range(160 if kind == "nightsky" else 60):
                sx, sy = rng.integers(0, w), rng.integers(0, int(h * .55))
                br = rng.uniform(.3, 1)
                base[sy:sy + 2, sx:sx + 2, :3] = base[sy:sy + 2, sx:sx + 2, :3] * (1 - br) + 255 * br
            img = Image.fromarray(base[..., :3].astype(np.uint8))
            dr = ImageDraw.Draw(img)
            if kind == "nightsky":
                far = [(0, .80), (.09, .72), (.16, .76), (.25, .66), (.34, .75), (.43, .69), (.52, .76), (.61, .64), (.70, .73), (.80, .67), (.90, .75), (1, .70), (1, 1), (0, 1)]
                near = [(0, .87), (.06, .82), (.11, .845), (.17, .79), (.22, .82), (.28, .765), (.33, .81), (.40, .785), (.46, .725), (.50, .66), (.54, .73), (.60, .78), (.66, .755), (.72, .805), (.78, .775), (.84, .82), (.90, .79), (.95, .83), (1, .81), (1, 1), (0, 1)]
                dr.polygon([(x * w, y * h) for x, y in far], fill=(64, 72, 160))
                mask = Image.new("L", (w, h), 0)
                ImageDraw.Draw(mask).polygon([(x * w, y * h) for x, y in near], fill=255)
                snow = np.zeros((h, w, 3), np.float32)
                tt = np.clip((np.linspace(0, 1, h) - .66) / .34, 0, 1)
                for c, (a, b) in enumerate(zip((233, 238, 252), (32, 35, 92))):
                    snow[..., c] = (a + (b - a) * tt)[:, None]
                img.paste(Image.fromarray(snow.astype(np.uint8)), (0, 0), mask)
                star = self._radial(w, h, w * .5, h * .30, 7, 7, [(0, (255, 255, 255), 1), (1, (255, 255, 255), 0), (9, (0, 0, 0), 0)])
                arr = self._over(np.dstack([np.asarray(img, np.float32), np.ones((h, w))]), star)
                return self._save(arr, "bg_" + kind)
            else:
                dr.polygon([(x * w, y * h) for x, y in [(0, .79), (.15, .77), (.30, .80), (.45, .76), (.60, .79), (.78, .755), (1, .785), (1, 1), (0, 1)]], fill=(4, 6, 12))
                return self._save(np.dstack([np.asarray(img, np.float32), np.ones((h, w))]), "bg_" + kind)
        elif kind == "holo":
            base[..., :3] = (233, 236, 255)
            for cx, cy, col in [(.15, .2, (168, 240, 255)), (.85, .25, (201, 182, 255)), (.7, .9, (255, 179, 230)), (.2, .9, (184, 255, 217))]:
                base = self._over(base, self._radial(w, h, w * cx, h * cy, w * .5, h * .6, [(0, col, 1), (1, col, 0), (9, col, 0)]))
        return self._save(base, "bg_" + kind)

    def orb(self):
        p = self._path("orb")
        if os.path.exists(p):
            return p
        s = 600
        g1, g2 = _hex(self.th["g1"]), _hex(self.th["g2"])
        a = self._radial(s, s, s * .76, s * .24, s * .95, s * .95,
                         [(0, _mix(g1, (255, 255, 255), .55), 1), (.2, g1, 1), (.45, g2, 1), (.72, (10, 3, 16), 1), (1, (0, 0, 0), 1), (9, (0, 0, 0), 1)])
        y, x = np.mgrid[0:s, 0:s]
        d = np.sqrt((x - s / 2) ** 2 + (y - s / 2) ** 2)
        a[..., 3] = np.clip((s / 2 - d) / 2.0, 0, 1)
        return self._save(a, "orb", alpha=True)

    def halo(self):
        p = self._path("halo")
        if os.path.exists(p):
            return p
        s = 800
        a = self._radial(s, s, s / 2, s / 2, s / 2, s / 2,
                         [(0, _hex(self.th["g1"]), .75), (.5, _hex(self.th["g3"]), .30), (1, _hex(self.th["g3"]), 0), (9, (0, 0, 0), 0)])
        return self._save(a, "halo", alpha=True)

    def sphere_black(self):
        p = self._path("sphere_black")
        if os.path.exists(p):
            return p
        s = 400
        a = self._radial(s, s, s * .34, s * .28, s * .8, s * .8, [(0, (140, 140, 140), 1), (.22, (42, 42, 42), 1), (.6, (7, 7, 7), 1), (1, (0, 0, 0), 1), (9, (0, 0, 0), 1)])
        y, x = np.mgrid[0:s, 0:s]
        d = np.sqrt((x - s / 2) ** 2 + (y - s / 2) ** 2)
        a[..., 3] = np.clip((s / 2 - 2 - d) / 2.0, 0, 1)
        return self._save(a, "sphere_black", alpha=True)

    def cube_white(self):
        p = self._path("cube_white")
        if os.path.exists(p):
            return p
        s = 360
        img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        sh = Image.new("L", (s, s), 0)
        ImageDraw.Draw(sh).rounded_rectangle((120, 150, 300, 310), 18, fill=70)
        sh = sh.filter(ImageFilter.GaussianBlur(26))
        img.paste((40, 40, 60, 255), (0, 0), sh)
        face = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        fd = ImageDraw.Draw(face)
        fd.rounded_rectangle((90, 90, 250, 250), 14, fill=(248, 248, 250, 255))
        face = face.rotate(-28, resample=Image.BICUBIC, center=(170, 170))
        img.alpha_composite(face)
        img.save(p)
        return p


# ─────────────────────────── 텍스트 마크업 ───────────────────────────
_TOK = re.compile(r"(\*\*.+?\*\*|==.+?==|~~.+?~~)")


def _runs(line):
    out = []
    for part in _TOK.split(line):
        if not part:
            continue
        if part.startswith("**"):
            out.append((part[2:-2], "b"))
        elif part.startswith("=="):
            out.append((part[2:-2], "a"))
        elif part.startswith("~~"):
            out.append((part[2:-2], "d"))
        else:
            out.append((part, "r"))
    return out


class Deck:
    def __init__(self, theme="pink", company="COMPANY NAME", asset_dir=None):
        self.th = THEMES[theme]
        self.theme = theme
        self.company = company
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = E(W), E(H)
        self.layouts = []
        self._asset_dir = asset_dir
        self.assets = None

    # ── 기본 ──
    def _slide(self, layout, bg=None, paper=False):
        if self.assets is None:
            self.assets = Assets(self._asset_dir or tempfile.mkdtemp(prefix="pl_assets_"), self.th)
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        if paper:
            s.background.fill.solid()
            s.background.fill.fore_color.rgb = RGBColor(255, 255, 255)
        else:
            s.background.fill.solid()
            s.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
            if bg:
                self.pic(s, self.assets.bg(bg), 0, 0, W, H)
        self.layouts.append(layout)
        s._pl_paper = paper
        return s

    def notes(self, s, text):
        s.notes_slide.notes_text_frame.text = text

    def pic(self, s, path, x, y, w, h):
        return s.shapes.add_picture(path, E(x), E(y), E(w), E(h))

    def _font(self, run, name, size_px, color, spc_em, bold=False):
        f = run.font
        f.size = Pt(size_px * .75)
        f.name = name
        f.bold = bold
        f.color.rgb = RGBColor.from_string(color)
        rPr = run._r.get_or_add_rPr()
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = rPr.makeelement(qn("a:ea"), {})
            rPr.append(ea)
        ea.set("typeface", name)
        rPr.set("spc", str(int(round(spc_em * size_px * .75 * 100))))

    def text(self, s, x, y, w, h, spec, size, *, kind="title", color=None, align="l", anchor="t",
             line=1.18, dot=True, weight="r", accent=None, spc=None, para_space=0):
        """spec 마크업 텍스트를 네이티브 텍스트 상자로. kind: title(Paperlogy) | body(Pretendard) | en(라벨)"""
        dark = not getattr(s, "_pl_paper", False)
        color = color or (WHITE if dark else INK)
        accent = accent or (self.th["on_dark"] if dark else self.th["accent"])
        dimc = "4A4B50" if dark else "B5B6BC"
        tb = s.shapes.add_textbox(E(x), E(y), E(w), E(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        lines = spec.split("\n")
        fam = F_TITLE if kind == "title" else F_BODY
        if spc is None:
            spc = -0.06 if kind == "title" else -0.03
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            p.line_spacing = line
            if para_space and i:
                p.space_before = Pt(para_space)
            runs = _runs(ln)
            # 마지막 줄 끝 '.'은 마침표 컬러
            tail_dot = dot and kind == "title" and i == len(lines) - 1 and runs and runs[-1][0].endswith(".")
            if tail_dot:
                t, st = runs[-1]
                runs[-1] = (t[:-1], st)
                runs.append((".", "dot"))
            for t, st in runs:
                if not t:
                    continue
                r = p.add_run()
                r.text = t
                if kind == "en":
                    self._font(r, F_EN, size, color if st != "a" else accent, 0.0)
                    continue
                wgt = "b" if st in ("b", "a") else weight
                if kind == "title":
                    name, bold = fam.get(wgt, fam["r"]), False
                else:
                    name, bold = (fam["b"], True) if wgt == "b" else (fam.get(wgt, fam["r"]), False)
                col = {"a": accent, "d": dimc, "dot": self.th["dot"] if dark else self.th["accent"]}.get(st, color)
                if st == "dot":
                    name = fam["b"] if kind == "title" else name
                self._font(r, name, size, col, spc, bold)
        return tb

    def label(self, s, x, y, w, txt, size=14, color=None, align="l", spacing=0.02):
        tb = self.text(s, x, y, w, size * 1.6, txt, size, kind="en", color=color or "CFCFD4", align=align)
        for r in tb.text_frame.paragraphs[0].runs:
            r._r.get_or_add_rPr().set("spc", str(int(spacing * size * .75 * 100)))
        return tb

    def rect(self, s, x, y, w, h, fill=None, alpha=None, line=None, line_alpha=None, line_w=1.0, radius=None, shape=None):
        shp = s.shapes.add_shape(shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE), E(x), E(y), E(w), E(h))
        if radius:
            shp.adjustments[0] = min(.5, radius / min(w, h))
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor.from_string(fill)
            if alpha is not None:
                self._alpha(shp.fill._xPr.find(qn("a:solidFill")), alpha)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = RGBColor.from_string(line)
            shp.line.width = Pt(line_w)
            if line_alpha is not None:
                self._alpha(shp.line._get_or_add_ln().find(qn("a:solidFill")), line_alpha)
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
        tf = shp.text_frame
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        return shp

    @staticmethod
    def _alpha(solid, a):
        clr = solid[0]
        for old in clr.findall(qn("a:alpha")):
            clr.remove(old)
        el = etree.SubElement(clr, qn("a:alpha"))
        el.set("val", str(int(a * 100000)))

    def glow(self, shp, color, alpha=.55, blur=18, dist=6):
        spPr = shp._element.spPr
        for old in spPr.findall(qn("a:effectLst")):
            spPr.remove(old)
        eff = etree.SubElement(spPr, qn("a:effectLst"))
        sh = etree.SubElement(eff, qn("a:outerShdw"), blurRad=str(int(blur * PX)), dist=str(int(dist * PX)), dir="5400000", algn="t", rotWithShape="0")
        c = etree.SubElement(sh, qn("a:srgbClr"), val=color)
        etree.SubElement(c, qn("a:alpha"), val=str(int(alpha * 100000)))

    def hline(self, s, x1, y1, x2, y2, color="FFFFFF", w=1.0, alpha=None):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2))
        ln.line.color.rgb = RGBColor.from_string(color)
        ln.line.width = Pt(w)
        if alpha is not None:
            self._alpha(ln.line._get_or_add_ln().find(qn("a:solidFill")), alpha)
        return ln

    def nails(self, s, x, y, w, h, kind="plus", inset=14, color="E9E9E9"):
        pts = [(x + inset, y + inset), (x + w - inset - 12, y + inset), (x + inset, y + h - inset - 12), (x + w - inset - 12, y + h - inset - 12)]
        if kind == "plus":
            for px_, py_ in pts:
                self.text(s, px_ - 2, py_ - 6, 16, 20, "+", 16, kind="en", color=color)
        elif kind == "corner":
            L = 12
            for i, (px_, py_) in enumerate(pts):
                xh = px_ if i % 2 == 0 else px_ + 12
                yv = py_ if i < 2 else py_ + 12
                self.hline(s, xh, yv, xh + (L if i % 2 == 0 else -L), yv, color, 1.5)
                self.hline(s, xh, yv, xh, yv + (L if i < 2 else -L), color, 1.5)
        elif kind == "circle":
            for px_, py_ in pts:
                self.rect(s, px_, py_, 12, 12, line=color, line_w=1.2, shape=MSO_SHAPE.OVAL)

    def tag(self, s, x, y, w, txt, color=None, h=30, size=13, glow=True):
        color = color or self.th["accent"]
        r = self.rect(s, x, y, w, h, fill=color)
        if glow:
            self.glow(r, color, .5, 16, 6)
        self.text(s, x, y, w, h, f"**{txt}**", size, kind="body", color="FFFFFF", align="c", anchor="m", line=1.0)
        return r

    def note(self, s, txt, y=678):
        self.text(s, 0, y, W, 18, txt, 12, kind="body", color="C4C4CC" if not s._pl_paper else "808080", align="c")

    def _rp_frame(self, s, head_l, head_r, page, dark=False):
        mute, strong = ("BBBBBB", "FFFFFF") if dark else ("777777", "333333")
        self.text(s, 58, 22, 600, 18, head_l, 12, kind="body", color=mute, accent=strong)
        self.text(s, 622, 22, 600, 18, head_r, 12, kind="body", color=mute, accent=strong, align="r")
        if dark:
            self.hline(s, 58, 52, W - 58, 52, "FFFFFF", .75, alpha=.25)
        else:
            self.hline(s, 58, 52, 640, 52, self.th["accent"], 1.0)
            self.hline(s, 640, 52, W - 58, 52, "C9C8F5", .75)
        self.text(s, 58, 690, 600, 16, self.company, 11, kind="en", color="DDDDDD" if dark else "808080")
        self.text(s, 1022, 690, 200, 16, str(page), 11, kind="en", color="DDDDDD" if dark else "808080", align="r")

    # ─────────────────────── 아키타입 ───────────────────────
    def cover(self, label, title, subtitle, meta, image=None, notes=None):
        """C01 표지. image=사진 경로(우측 절반, 좌측 그라데이션 마스크). 없으면 후광+구체."""
        s = self._slide("C01", "aurora_bottom")
        if image:
            self.pic(s, image, 560, 0, 720, 720)
            self._mask_left(s, 560, 720)
        else:
            self.pic(s, self.assets.halo(), 640, -60, 760, 760)
            self.pic(s, self.assets.orb(), 820, 170, 380, 380)
        self.label(s, 96, 236, 600, label)
        self.text(s, 96, 266, 760, 150, title, 54, line=1.15)
        self.text(s, 96, 424, 760, 30, subtitle, 18, kind="body", color=BODY)
        self.text(s, 96, 560, 600, 50, meta, 14, kind="body", color="CFCFD4", accent=WHITE, line=1.5)
        if notes:
            self.notes(s, notes)
        return s

    def _mask_left(self, s, x, w):
        """사진 왼쪽을 검정으로 페이드(텍스트 영역 확보)"""
        shp = self.rect(s, x, 0, w, H, fill="000000")
        sp = shp.fill._xPr
        sf = sp.find(qn("a:solidFill"))
        sp.remove(sf)
        grad = etree.SubElement(sp, qn("a:gradFill"), rotWithShape="1")
        gl = etree.SubElement(grad, qn("a:gsLst"))
        for pos, a in ((0, 100), (45, 55), (75, 0)):
            g = etree.SubElement(gl, qn("a:gs"), pos=str(pos * 1000))
            c = etree.SubElement(g, qn("a:srgbClr"), val="000000")
            etree.SubElement(c, qn("a:alpha"), val=str(a * 1000))
        etree.SubElement(grad, qn("a:lin"), ang="0", scaled="0")
        # gradFill은 prstGeom 뒤, ln 앞이어야 함
        ln = sp.find(qn("a:ln"))
        if ln is not None:
            sp.remove(grad); ln.addprevious(grad)
        return shp

    def photo_text(self, title, body, image, notes=None):
        """K07 사진 + 좌측 텍스트 (풀블리드 사진, 좌측 마스크)"""
        s = self._slide("K07")
        self.pic(s, image, 0, 0, W, H)
        self._mask_left(s, 0, W)
        self.text(s, 78, 250, 560, 130, title, 46, line=1.2)
        self.text(s, 78, 390, 520, 100, body, 17, kind="body", color=BODY, accent=WHITE, line=1.55)
        if notes:
            self.notes(s, notes)
        return s

    def toc(self, items, active=0, title="오늘\n이야기할 것.", notes=None):
        """C03 목차. items=[(제목, 설명)], active=현재 챕터 index"""
        s = self._slide("C03", "aurora_bottom")
        self.label(s, 96, 250, 380, "CONTENTS")
        self.text(s, 96, 276, 400, 130, "\n".join(f"**{ln}**" for ln in title.split("\n")), 48, line=1.18)
        n = len(items)
        gap = min(118, 470 // max(n, 1))
        top = 360 - gap * n / 2
        for i, (t, d) in enumerate(items):
            on = i == active
            c = WHITE if on else "4A4B50"
            self.text(s, 560, top + i * gap, 110, 64, f"{i + 1:02d}", 52, kind="en", color=c)
            self.text(s, 680, top + i * gap + 2, 520, 40, f"**{t}**", 28, color=c, dot=False)
            self.text(s, 680, top + i * gap + 42, 520, 24, d, 15, kind="body", color=BODY if on else "4A4B50")
        if notes:
            self.notes(s, notes)
        return s

    def section(self, label, title, notes=None):
        """K01 섹션 구분(야경)"""
        s = self._slide("K01", "nightsky")
        self.label(s, 0, 286, W, label, align="c", color="FFFFFF")
        self.text(s, 0, 312, W, 60, title, 40, align="c", dot=False)
        if notes:
            self.notes(s, notes)
        return s

    def statement(self, text, note=None, size=58, notes=None):
        """K02 스테이트먼트 / K09 매니페스토"""
        s = self._slide("K02", "aurora_top")
        self.text(s, 80, 250, W - 160, 220, text, size, align="c", anchor="m", line=1.2)
        if note:
            self.note(s, note)
        if notes:
            self.notes(s, notes)
        return s

    def manifesto(self, text, notes=None):
        s = self.statement(text, size=68, notes=notes)
        self.layouts[-1] = "K09"
        return s

    def principle_card(self, title, body, before, after, nail="plus", notes=None):
        """K03 원칙 + 글래스 카드"""
        s = self._slide("K03", "nebula")
        self.text(s, 120, 222, 560, 140, title, 50, line=1.18)
        self.text(s, 120, 380, 540, 120, body, 19, kind="body", color=BODY, accent=WHITE, line=1.55)
        x, y, w, h = 720, 206, 440, 308
        self.rect(s, x, y, w, h, fill="121418", alpha=.80, line="FFFFFF", line_alpha=.10, radius=8)
        self.nails(s, x, y, w, h, nail)
        self.text(s, x + 30, y + 52, w - 60, 66, before, 21, color="8A8C93", align="c", line=1.3, dot=False)
        self.text(s, x + 30, y + 128, w - 60, 26, "그래서  ↓", 15, kind="body", color=WHITE, align="c")
        self.text(s, x + 30, y + 166, w - 60, 100, after, 24, color=self.th["on_dark"] if self.theme != "pink" else "37B68F", align="c", line=1.3, dot=False, weight="b")
        if notes:
            self.notes(s, notes)
        return s

    def word_stack(self, title, body, words, active, notes=None):
        """K04 워드 스택 순차 공개 (active 위치만 흰색). 같은 내용으로 active만 바꿔 3장 만든다."""
        s = self._slide("K04", "aurora_bottom")
        self.text(s, 96, 236, 560, 140, title, 50, line=1.18)
        self.text(s, 96, 392, 560, 110, body, 20, kind="body", color=BODY, accent=WHITE, line=1.5)
        for i, wd in enumerate(words):
            self.text(s, 690, 150 + i * 128, 560, 150, wd, 112, kind="en", color=WHITE if i == active else "333333")
        if notes:
            self.notes(s, notes)
        return s

    def orb_text(self, title, text, note=None, notes=None):
        """K05 오브 센터"""
        s = self._slide("K05", "aurora_top")
        self.text(s, 0, 118, W, 60, title, 40, align="c", dot=False)
        self.pic(s, self.assets.orb(), 530, 250, 220, 220)
        self.text(s, 140, 300, W - 280, 120, text, 22, kind="body", align="c", anchor="m", line=1.5, weight="b")
        if note:
            self.note(s, note)
        if notes:
            self.notes(s, notes)
        return s

    def number_hero(self, title, number, unit, evidence, note=None, notes=None):
        """K06 숫자 히어로. 숫자 글자 수에 따라 크기·근거 위치 자동 조정"""
        s = self._slide("K06", "aurora_top")
        self.text(s, 0, 118, W, 60, title, 40, align="c", dot=False)
        n = len(number) + len(unit) * .55
        size = 210 if n <= 3.5 else 170 if n <= 5 else 130
        num_w = size * .62 * n
        total = num_w + 110 + 340
        x0 = max(60, (W - total) / 2)
        self.pic(s, self.assets.orb(), x0 + num_w - 150, 250, 230, 230)
        tb = self.text(s, x0, 365 - size * .72, num_w + 40, size * 1.44, number, size, kind="en", color=WHITE, anchor="m")
        if unit:
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = unit
            self._font(r, F_EN, size * .45, WHITE, 0)
        self.text(s, x0 + num_w + 110, 300, 380, 140, evidence, 20, kind="body", accent=WHITE, color=BODY, line=1.5, anchor="m")
        if note:
            self.note(s, note)
        if notes:
            self.notes(s, notes)
        return s

    def flow(self, title, sub, steps, emotions=None, notes=None):
        """K10 흐름 원 (3–4단계)"""
        s = self._slide("K10", "aurora_top")
        self.text(s, 0, 140, W, 60, title, 40, align="c", dot=False)
        self.text(s, 0, 206, W, 30, sub, 18, kind="body", color=BODY, align="c")
        n = len(steps)
        d, gap = 150, 110
        tot = n * d + (n - 1) * gap
        x0 = (W - tot) / 2
        for i, st in enumerate(steps):
            x = x0 + i * (d + gap)
            c = self.rect(s, x, 300, d, d, fill="17171B", line="FFFFFF", line_alpha=.18, line_w=1.5, shape=MSO_SHAPE.OVAL)
            self.glow(c, self.th["g1"], .65, 26, 0)
            self.text(s, x + 10, 300, d - 20, d, st, 20, align="c", anchor="m", line=1.2, weight="b", dot=False)
            if i < n - 1:
                self.hline(s, x + d + 8, 375, x + d + gap - 8, 375, "FFFFFF", 1.5, .8)
            if emotions:
                self.text(s, x - 20, 472, d + 40, 26, emotions[i], 15, kind="body", align="c", weight="b")
        if notes:
            self.notes(s, notes)
        return s

    def timeline(self, title, sub, items, notes=None):
        """K14 세로 타임라인. items=[(연도, 내용, 강조bool)]"""
        s = self._slide("K14", "dark_report")
        self.text(s, 110, 290, 380, 110, title, 36, align="c", line=1.25, weight="b")
        self.text(s, 110, 404, 380, 24, sub, 14, kind="body", color="AAAAAA", align="c")
        self.hline(s, 560, 70, 560, 650, "FFFFFF", .75, .7)
        n = len(items)
        step = min(86, 520 / max(n, 1))
        top = 360 - step * (n - 1) / 2
        for i, (yr, txt, hl) in enumerate(items):
            y = top + i * step
            c = self.th["on_dark"] if hl else WHITE
            self.rect(s, 556, y - 4, 9, 9, fill=c, shape=MSO_SHAPE.OVAL)
            self.text(s, 590, y - 14, 120, 28, yr, 17, kind="en", color=c, anchor="m")
            self.text(s, 700, y - 14, 520, 28, (f"**{txt}**" if hl else txt), 17, kind="body", color=c, anchor="m")
        if notes:
            self.notes(s, notes)
        return s

    def big_med_small(self, title, items, notes=None):
        """K15 빅·미디엄·스몰. items=[(라벨, 설명)] 3개 (중요도 순)"""
        s = self._slide("K15")
        self.text(s, 96, 100, 1000, 50, title, 32, line=1.2)
        geo = [(300, 200, 300), (650, 190, 170), (620, 400, 110)]
        for (lab, _), (x, y, d) in zip(items, geo):
            self.rect(s, x, y, d, d, line="FFFFFF", line_w=1.5, shape=MSO_SHAPE.OVAL)
            self.text(s, x, y, d, d, f"**{lab}**", max(20, int(d * .16)), align="c", anchor="m", dot=False, line=1.15)
        self.text(s, 880, 250, 360, 200, "\n".join(f"**{l}**  {t}" for l, t in items), 17, kind="body", color=BODY, accent=WHITE, line=1.8)
        if notes:
            self.notes(s, notes)
        return s

    def ending(self, title, request, notes=None):
        """C02 엔딩/요청"""
        s = self._slide("C02", "nightsky")
        self.text(s, 0, 240, W, 70, title, 48, align="c")
        self.text(s, 0, 318, W, 30, request, 18, kind="body", color="E6E6EE", accent=WHITE, align="c")
        if notes:
            self.notes(s, notes)
        return s

    # ── 보고형 ──
    def _left_title(self, s, title, desc, dark=False, size=33):
        self.text(s, 58, 92, 310, 200, title, size, line=1.24, color=WHITE if dark else INK)
        self.text(s, 58, 300, 300, 90, desc, 14, kind="body", color="DDDDDD" if dark else INK_BODY, accent=WHITE if dark else INK, line=1.55)

    def report_cards(self, head_l, head_r, page, title, desc, headline, cards, table=None, conclusion=None, notes=None):
        """R01 라이트: cards=[(태그, 숫자, 단위, 미터%, 설명, 색키 'a'|'b'|'g')], table=[[헤더..],[행..],...]"""
        s = self._slide("R01", paper=True)
        self._rp_frame(s, head_l, head_r, page)
        self._left_title(s, title, desc)
        self.pic(s, self.assets.cube_white(), 70, 470, 200, 200)
        self.text(s, 400, 92, 822, 36, headline, 24, align="c", dot=False)
        n = len(cards)
        gap = 18
        cw = (822 - gap * (n - 1)) / n
        col = {"a": self.th["accent"], "b": self.th["accent2"], "g": "5A5B60"}
        card_h = 270 if table else 400
        for i, (tg, num, unit, pct, dsc, ck) in enumerate(cards):
            x = 400 + i * (cw + gap)
            c = col.get(ck, ck)
            self.tag(s, x, 146, cw, tg, c if ck != "g" else "4A4B50")
            self.rect(s, x, 194, cw, card_h, line=LINE, line_w=.75, radius=4)
            tb = self.text(s, x + 18, 208, cw - 36, 80, num, 64, kind="en", color=c)
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = unit
            self._font(r, F_EN, 30, c, 0)
            self.text(s, x + 18, 300, cw - 36, 16, "응답 비중" if pct is not None else "", 11, kind="body", color="8A8A8A")
            if pct is not None:
                self.rect(s, x + 18, 320, cw - 36, 10, fill=PAPER2, radius=5)
                self.rect(s, x + 18, 320, (cw - 36) * pct / 100, 10, fill=c, radius=5)
            self.text(s, x + 18, 194 + card_h - 76, cw - 36, 62, dsc, 13.5, kind="body", color=INK_BODY, accent=INK, line=1.5, anchor="b")
        if table:
            ty, th_ = 194 + card_h + 18, 620 - (194 + card_h + 18)
            self.table(s, 400, ty, 822, th_, table)
        if conclusion:
            b = self.rect(s, 400, 632, 822, 38, line=self.th["accent"], line_w=1.5, radius=4)
            self.text(s, 400, 632, 822, 38, f"**{conclusion}**", 15, kind="body", color=self.th["accent"], align="c", anchor="m")
        if notes:
            self.notes(s, notes)
        return s

    def table(self, s, x, y, w, h, rows, dark=False, hl_col=None):
        nr, nc = len(rows), len(rows[0])
        gt = s.shapes.add_table(nr, nc, E(x), E(y), E(w), E(h))
        tbl = gt.table
        tblPr = tbl._tbl.tblPr
        for k in ("firstRow", "bandRow"):
            tblPr.set(k, "0")
        style = tblPr.find(qn("a:tableStyleId"))
        if style is None:
            style = etree.SubElement(tblPr, qn("a:tableStyleId"))
        style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"   # No Style, Table Grid
        for r in range(nr):
            tbl.rows[r].height = E(h / nr)
            for c in range(nc):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = E(8)
                cell.margin_top = cell.margin_bottom = E(2)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid()
                if r == 0:
                    cell.fill.fore_color.rgb = RGBColor.from_string("3A3A3A" if dark else PAPER2)
                else:
                    cell.fill.fore_color.rgb = RGBColor.from_string("1F1F1F" if dark else "FFFFFF")
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                txt = str(rows[r][c])
                for t, st in _runs(txt):
                    run = p.add_run()
                    run.text = t
                    bold = r == 0 or st in ("b", "a")
                    colr = (WHITE if dark else "333333") if r == 0 else (("CFCFCF" if dark else "333333") if st == "r" else (self.th["on_dark"] if dark else self.th["accent"]) if st == "a" else (WHITE if dark else INK))
                    self._font(run, "Pretendard", 13, colr, -0.02, bold)
                tcPr = cell._tc.get_or_add_tcPr()
                for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                    ln = etree.SubElement(tcPr, qn(side), w=str(int(.75 * 12700)))
                    sf = etree.SubElement(ln, qn("a:solidFill"))
                    etree.SubElement(sf, qn("a:srgbClr"), val="4A4A4A" if dark else LINE)
                # 스키마 순서: lnX → fill. fill이 먼저 붙어 있으면 뒤로 이동
                fill = tcPr.find(qn("a:solidFill"))
                if fill is not None:
                    tcPr.remove(fill)
                    tcPr.append(fill)
        return gt

    def report_kpi_chart(self, head_l, head_r, page, title, desc, kpis, chart_title, cats, vals, roadmap_title, roadmap, conclusion, note=None, number_format='0.0', notes=None):
        """R04 KPI 3 + 네이티브 막대차트(마지막 막대 강조) + 로드맵. roadmap=[(단계, 설명)] 첫 항목 활성"""
        s = self._slide("R04", paper=True)
        self._rp_frame(s, head_l, head_r, page)
        self._left_title(s, title, desc)
        self.pic(s, self.assets.cube_white(), 70, 470, 200, 200)
        kw = (822 - 36) / 3
        for i, (num, lab, acc) in enumerate(kpis):
            x = 400 + i * (kw + 18)
            self.rect(s, x, 92, kw, 96, fill=PAPER2, radius=6)
            self.text(s, x + 18, 104, kw - 36, 50, num, 40, kind="en", color=self.th["accent"] if acc else INK)
            self.text(s, x + 18, 156, kw - 36, 20, lab, 13, kind="body", color=INK_BODY)
        self.text(s, 400, 208, 420, 24, chart_title, 15, kind="body", color=INK_BODY, accent=INK)
        cd = CategoryChartData()
        cd.categories = cats
        cd.add_series("값", vals)
        gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, E(400), E(236), E(420), E(372), cd)
        ch = gf.chart
        ch.has_legend = False
        ch.has_title = False
        ch.font.size = Pt(9.5)
        ch.font.name = "Pretendard"
        ch.font.color.rgb = RGBColor.from_string("595959")
        va = ch.value_axis
        va.visible = False
        va.has_major_gridlines = False
        ca = ch.category_axis
        ca.format.line.color.rgb = RGBColor.from_string(LINE)
        ca.tick_labels.font.size = Pt(10)
        pl = ch.plots[0]
        pl.gap_width = 55
        pl.has_data_labels = True
        dl = pl.data_labels
        dl.number_format = number_format
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.font.size = Pt(10)
        ser = pl.series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string("C9CAD3")
        for k in range(len(vals)):
            q = ser.points[k].data_label
            q.position = XL_LABEL_POSITION.OUTSIDE_END
            q.number_format = number_format
            q.number_format_is_linked = False
            q.show_value = True
            q.font.size = Pt(10)
            q.font.color.rgb = RGBColor.from_string("595959")
        pt = ser.points[len(vals) - 1]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = RGBColor.from_string(self.th["accent"])
        pt.data_label.font.bold = True
        pt.data_label.font.color.rgb = RGBColor.from_string(self.th["accent"])
        self.text(s, 850, 208, 372, 24, roadmap_title, 15, kind="body", color=INK_BODY, accent=INK)
        n = len(roadmap)
        step = 360 / n
        for i, (st, txt) in enumerate(roadmap):
            y = 246 + i * step + (step - 30) / 2
            self.rect(s, 850, y, 86, 30, fill=self.th["accent"] if i == 0 else "5A5B60", radius=15)
            self.text(s, 850, y, 86, 30, f"**{st}**", 12, kind="body", color="FFFFFF", align="c", anchor="m")
            self.text(s, 948, y - 8, 274, 46, txt, 13.5, kind="body", color=INK_BODY, accent=INK, anchor="m", line=1.35)
        self.rect(s, 400, 632, 822, 38, line=self.th["accent"], line_w=1.5, radius=4)
        self.text(s, 400, 632, 822, 38, f"**{conclusion}**", 15, kind="body", color=self.th["accent"], align="c", anchor="m")
        if note:
            self.text(s, 400, 610, 822, 16, note, 11, kind="body", color="808080", align="r")
        if notes:
            self.notes(s, notes)
        return s

    def report_dark_table(self, head_l, head_r, page, title, desc, headline, rows, notes=None):
        """R03 다크 매트릭스(표)"""
        s = self._slide("R03", "dark_report")
        self.rect(s, 0, 150, W, 440, fill="1F1F1F")
        self._rp_frame(s, head_l, head_r, page, dark=True)
        self._left_title(s, title, desc, dark=True, size=36)
        self.pic(s, self.assets.sphere_black(), 100, 450, 140, 140)
        self.text(s, 400, 92, 822, 36, headline, 22, align="c", dot=False)
        self.table(s, 400, 150, 822, 440, rows, dark=True)
        if notes:
            self.notes(s, notes)
        return s

    # ── 저장 ──
    def save(self, path):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.prs.save(path)
        with open(os.path.splitext(path)[0] + ".layouts.json", "w", encoding="utf-8") as f:
            json.dump(self.layouts, f)
        return path

