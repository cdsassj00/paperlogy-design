# -*- coding: utf-8 -*-
"""
render_pptx.py — PowerPoint(COM)로 PPTX를 실제 렌더해 PNG·컨택트 시트를 만들고 검사한다.

    python render_pptx.py deck.pptx out_dir            # PNG + contact.png
    python render_pptx.py deck.pptx out_dir --check    # + 검사

검사(--check) — PowerPoint가 계산한 실제 글자 경계(TextRange.Bound*)를 쓴다.
  FAIL  글자 < 11px(8.25pt) / 텍스트 박스 넘침(글자 높이 > 상자 높이) / 캔버스 가장자리 침범 /
        서로 다른 텍스트 상자의 글자 겹침 / 설치되지 않은 폰트 사용(대체 글꼴로 깨짐) /
        R*(보고형) slide-density 채움 미달 (deck.layouts.json 필요 — pptx_builder가 생성)
  WARN  판독 위험 대비(큰 글자 2.5:1·일반 3:1 미만)
Windows + PowerPoint 필요. 종료코드: FAIL 있으면 1.
"""
import sys, os, json, glob, pathlib, subprocess, statistics
from PIL import Image, ImageDraw, ImageFont
import win32com.client
from fontTools.ttLib import TTFont

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

CHECK_FILL = pathlib.Path.home() / ".claude/skills/slide-density/scripts/check_fill.py"
K = 1280 / 960  # pt → px


def installed_families():
    fams = set()
    dirs = [r"C:\Windows\Fonts", os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Windows\Fonts")]
    for d in dirs:
        for f in glob.glob(os.path.join(d, "*.tt[fc]")) + glob.glob(os.path.join(d, "*.otf")):
            try:
                fonts = TTFont(f, fontNumber=0, lazy=True)
                for rec in fonts["name"].names:
                    if rec.nameID in (1, 4, 16):
                        fams.add(rec.toUnicode().strip().lower())
            except Exception:
                pass
    return fams


def _lum(c):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * ch(c[0]) + 0.7152 * ch(c[1]) + 0.0722 * ch(c[2])


def _font(size):
    for n in ("malgunbd.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(n, size)
        except Exception:
            pass
    return ImageFont.load_default()


def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(2)
    src = os.path.abspath(sys.argv[1]); out = pathlib.Path(sys.argv[2]).absolute()
    out.mkdir(parents=True, exist_ok=True)
    check = "--check" in sys.argv
    lay_path = os.path.splitext(src)[0] + ".layouts.json"
    layouts = json.load(open(lay_path, encoding="utf-8")) if os.path.exists(lay_path) else []

    app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")  # 조기 바인딩(Runs() 등 메서드 정상 호출)
    pres = app.Presentations.Open(src, True, False, False)  # ReadOnly, Untitled=False, WithWindow=False
    sw, sh = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
    info, used_fonts = [], set()
    try:
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)
            lay = layouts[i - 1] if i - 1 < len(layouts) else "X"
            png = out / f"{i:02d}_{lay}.png"
            sl.Export(str(png), "PNG", 1280, int(1280 * sh / sw))
            texts = []
            for shp in sl.Shapes:
                if not shp.HasTextFrame or not shp.TextFrame.HasText:
                    continue
                tr = shp.TextFrame.TextRange   # 레거시 TextRange(조기 바인딩에서 Runs(i,1) 동작)
                txt = " ".join(tr.Text.split())[:16]
                sizes, cols = [], []
                for r in range(1, 200):
                    run = tr.Runs(r, 1)
                    if not run.Text:
                        break
                    if not run.Text.strip():
                        continue
                    sizes.append(run.Font.Size)
                    used_fonts.add(run.Font.Name)
                    used_fonts.add(run.Font.NameFarEast)
                    bgr = run.Font.Color.RGB
                    cols.append((bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255))
                texts.append(dict(t=txt, box=(shp.Left, shp.Top, shp.Width, shp.Height),
                                  bound=(tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight),
                                  size=min(sizes) if sizes else 0, maxsize=max(sizes) if sizes else 0, color=cols[0] if cols else (255, 255, 255)))
            info.append((png, lay, texts))
    finally:
        pres.Close()
        if app.Presentations.Count == 0:
            app.Quit()

    # contact sheet
    w, cols_n = 480, 4
    h = w * 9 // 16
    rows = (len(info) + cols_n - 1) // cols_n
    sheet = Image.new("RGB", (cols_n * w, rows * (h + 26)), "white")
    d, fnt = ImageDraw.Draw(sheet), _font(16)
    for k, (png, lay, _) in enumerate(info):
        x, y = (k % cols_n) * w, (k // cols_n) * (h + 26)
        sheet.paste(Image.open(png).resize((w, h)), (x, y + 26))
        d.text((x + 6, y + 4), png.stem, fill="red", font=fnt)
    sheet.save(out / "contact.png")
    print(f"렌더 {len(info)}장 → {out}\n컨택트 시트: {out / 'contact.png'}")
    if not check:
        return

    fail = 0
    fams = installed_families()
    missing = sorted(f for f in used_fonts if f and not f.startswith("+") and f.lower() not in fams)
    for f in missing:
        print(f"[FAIL] 설치되지 않은 폰트: {f}  → assets/fonts/ttf 설치 또는 이름 확인")
    fail += bool(missing)
    for png, lay, texts in info:
        bad, warn = [], []
        im = Image.open(png).convert("RGB"); px = im.load()
        for t in texts:
            bl, bt, bw, bh = t["bound"]; x, y, ww, hh = t["box"]
            if t["size"] and t["size"] < 8.25:
                bad.append(f"글자<11px: {t['t']}")
            if bh > hh + 3:
                bad.append(f"상자 넘침(+{(bh - hh) * K:.0f}px): {t['t']}")
            if bl * K < 15 or (bl + bw) * K > 1265 or bt * K < 8 or (bt + bh) * K > 712:
                bad.append(f"가장자리 침범: {t['t']}")
            # contrast
            fg = t["color"]; bgs = []
            for yy in range(int(bt * K), int((bt + bh) * K), 3):
                for xx in range(int(bl * K), int((bl + bw) * K), 3):
                    if 0 <= xx < im.width and 0 <= yy < im.height:
                        c = px[xx, yy]
                        if sum(abs(a - b) for a, b in zip(c, fg)) > 90:
                            bgs.append(_lum(c))
            if len(bgs) > 8:
                b, f = statistics.median(bgs), _lum(fg)
                ratio = (max(b, f) + .05) / (min(b, f) + .05)
                need = 2.5 if t["maxsize"] >= 18 else 3.0
                if ratio < need and fg not in ((51, 51, 51), (74, 75, 80), (58, 59, 64)):
                    warn.append(f"대비 {ratio:.1f}:1: {t['t']}")
        for a in range(len(texts)):
            for b in range(a + 1, len(texts)):
                A, B = texts[a]["bound"], texts[b]["bound"]
                ix = min(A[0] + A[2], B[0] + B[2]) - max(A[0], B[0])
                iy = min(A[1] + A[3], B[1] + B[3]) - max(A[1], B[1])
                if ix > 3 and iy > 0.3 * min(A[3], B[3]):
                    bad.append(f"텍스트 겹침: {texts[a]['t']} ↔ {texts[b]['t']}")
        if lay.startswith("R") and CHECK_FILL.exists():
            r = subprocess.run([sys.executable, str(CHECK_FILL), str(png)], capture_output=True, text=True, encoding="utf-8", errors="replace")
            if r.returncode:
                bad.append("채움 미달(slide-density)\n" + r.stdout.strip()[-500:])
        for m in bad:
            print(f"[FAIL] {png.name} {m}")
        for m in warn:
            print(f"[WARN] {png.name} {m}")
        tag = "보고형" if lay.startswith("R") else "키노트 — contact.png 육안 검토 필요"
        print(f"[{'FAIL' if bad else 'PASS'}] {png.name} ({tag})")
        fail += bool(bad)
    sys.exit(1 if fail else 0)


if __name__ == "__main__":
    main()
