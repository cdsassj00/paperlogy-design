# -*- coding: utf-8 -*-
"""
render_slides.py — Paperlogy HTML 덱의 각 .slide를 1280x720 PNG로 렌더하고
컨택트 시트와 자동 검사를 수행한다.

사용:
    python render_slides.py deck.html out_dir            # PNG + contact.png
    python render_slides.py deck.html out_dir --check    # + 자동 검사

검사(--check):
  FAIL  모든 장표: 글자 11px 미만 / 텍스트가 캔버스 가장자리(좌우 20px·상하 12px) 침범 /
        서로 다른 텍스트 블록끼리 겹침·4px 미만으로 붙음 / 높이 고정 박스 밖으로 텍스트 넘침
  FAIL  R*(보고형): slide-density check_fill (빈 띠 ≤12%)
  WARN  판독 위험 대비(텍스트색 vs 실제 렌더 배경, 큰 글자 2.5:1·일반 3:1 미만).
        의도적으로 죽인 회색 텍스트는 class="c-dim" 또는 data-dim 속성으로 제외.
  K*(키노트)는 의도적 여백이라 채움 검사를 하지 않는다 → contact.png 육안 검토 필수.
종료코드: FAIL이 하나라도 있으면 1.

Windows 긴 경로(260자 초과)에서도 동작하도록 HTML 폴더를 로컬 HTTP 서버로 띄워 렌더한다.
"""
import sys, os, subprocess, pathlib, threading, functools, statistics, re
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from playwright.sync_api import sync_playwright
from PIL import Image, ImageDraw, ImageFont

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

CHECK_FILL = pathlib.Path.home() / ".claude/skills/slide-density/scripts/check_fill.py"

DOM_CHECK = r"""(slide) => {
  const S = slide.getBoundingClientRect(), bad = [], texts = [];
  const blockOf = el => { while (el && el !== slide) { const d = getComputedStyle(el).display;
      if (!d.startsWith('inline')) return el; el = el.parentElement; } return slide; };
  const fixedBox = el => { while (el && el !== slide) { if (el.style.height) return el; el = el.parentElement; } return null; };
  const w = document.createTreeWalker(slide, NodeFilter.SHOW_TEXT);
  while (w.nextNode()) {
    const n = w.currentNode; if (!n.textContent.trim()) continue;
    const el = n.parentElement, cs = getComputedStyle(el);
    if (cs.visibility === 'hidden' || cs.display === 'none') continue;
    const rg = document.createRange(); rg.selectNodeContents(n);
    const rects = [...rg.getClientRects()].filter(b => b.width > 1 && b.height > 1);
    if (!rects.length) continue;
    const t = n.textContent.trim().slice(0, 14), fs = parseFloat(cs.fontSize);
    if (fs < 11) bad.push('글자<11px: ' + t);
    const box = fixedBox(el), bb = box && box.getBoundingClientRect();
    for (const b of rects) {
      if (b.left < S.left + 20 || b.right > S.right - 20 || b.top < S.top + 12 || b.bottom > S.bottom - 12) { bad.push('가장자리 침범: ' + t); break; }
      if (bb && (b.bottom > bb.bottom + 2 || b.right > bb.right + 2)) { bad.push('박스 넘침: ' + t); break; }
    }
    const dim = !!el.closest('.c-dim,[data-dim]');
    texts.push({t, blk: blockOf(el), fs, bold: parseInt(cs.fontWeight) >= 700, color: cs.color, dim,
      rects: rects.map(b => [b.left - S.left, b.top - S.top, b.right - S.left, b.bottom - S.top])});
  }
  for (let i = 0; i < texts.length; i++) for (let j = i + 1; j < texts.length; j++) {
    const A = texts[i], B = texts[j];
    if (A.blk === B.blk || A.blk.contains(B.blk) || B.blk.contains(A.blk)) continue;
    for (const a of A.rects) for (const b of B.rects) {
      const ix = Math.min(a[2], b[2]) - Math.max(a[0], b[0]), iy = Math.min(a[3], b[3]) - Math.max(a[1], b[1]);
      const hmin = Math.min(a[3] - a[1], b[3] - b[1]);
      if (iy > 0.35 * hmin && ix > -4) { bad.push((ix > 0 ? '텍스트 겹침: ' : '텍스트 붙음(<4px): ') + A.t + ' ↔ ' + B.t); a[2] = -1e9; }
    }
  }
  return {bad: [...new Set(bad)], texts: texts.map(x => ({t: x.t, fs: x.fs, bold: x.bold, color: x.color, dim: x.dim, rects: x.rects}))};
}"""


def serve(folder):
    """긴 경로 대응: 폴더를 127.0.0.1 임의 포트로 서빙."""
    d = str(pathlib.Path(folder).absolute())
    if os.name == "nt" and len(d) > 200 and not d.startswith("\\\\?\\"):
        d = "\\\\?\\" + d
    class Quiet(SimpleHTTPRequestHandler):
        def log_message(self, *a):
            pass
    srv = ThreadingHTTPServer(("127.0.0.1", 0), functools.partial(Quiet, directory=d))
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


def render(html, out):
    out.mkdir(parents=True, exist_ok=True)
    html = pathlib.Path(html).absolute()
    srv = serve(html.parent)
    url = f"http://127.0.0.1:{srv.server_address[1]}/{html.name}?all=1"
    files = []
    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(viewport={"width": 1280, "height": 720}, device_scale_factor=1)
            pg.goto(url, wait_until="networkidle")
            pg.evaluate("document.fonts.ready")
            pg.wait_for_timeout(400)
            for i, el in enumerate(pg.query_selector_all(".slide"), 1):
                lay = el.get_attribute("data-layout") or "X"
                f = out / f"{i:02d}_{lay}.png"
                el.screenshot(path=str(f))
                files.append((f, lay, el.evaluate(DOM_CHECK)))
            b.close()
    finally:
        srv.shutdown()
    if not files:
        print(f"[ERROR] .slide를 하나도 찾지 못함 — 경로/파일 확인: {html}"); sys.exit(1)
    return files


def _font(size):
    for n in ("malgunbd.ttf", "malgun.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(n, size)
        except Exception:
            pass
    return ImageFont.load_default()


def contact(files, out, cols=4, w=480):
    h = w * 9 // 16
    rows = (len(files) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * w, rows * (h + 26)), "white")
    d, fnt = ImageDraw.Draw(sheet), _font(16)
    for k, (f, lay, _) in enumerate(files):
        im = Image.open(f).resize((w, h))
        x, y = (k % cols) * w, (k // cols) * (h + 26)
        sheet.paste(im, (x, y + 26))
        d.text((x + 6, y + 4), f.stem, fill="red", font=fnt)
    p = out / "contact.png"
    sheet.save(p)
    return p


def _rgb(css):
    v = [float(x) for x in re.findall(r"[\d.]+", css)[:3]]
    return tuple(int(c) for c in v) if len(v) == 3 else (255, 255, 255)


def _lum(c):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = c
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast_warnings(png, texts):
    im = Image.open(png).convert("RGB")
    px, W, H = im.load(), im.width, im.height
    warns = []
    for t in texts:
        if t["dim"]:
            continue
        fg = _rgb(t["color"])
        bgs = []
        for x0, y0, x1, y1 in t["rects"]:
            for y in range(max(0, int(y0)), min(H, int(y1)), 2):
                for x in range(max(0, int(x0)), min(W, int(x1)), 2):
                    c = px[x, y]
                    if sum(abs(a - b) for a, b in zip(c, fg)) > 90:  # 글자 획이 아닌 픽셀 = 배경
                        bgs.append(_lum(c))
        if len(bgs) < 8:
            continue
        bl, fl = statistics.median(bgs), _lum(fg)
        ratio = (max(bl, fl) + 0.05) / (min(bl, fl) + 0.05)
        need = 2.5 if (t["fs"] >= 24 or (t["fs"] >= 18.6 and t["bold"])) else 3.0  # 디자인 회색 보조텍스트(3:1대)는 허용, 판독 불가만 경고
        if ratio < need:
            warns.append(f"대비 {ratio:.1f}:1 < {need}: {t['t']}")
    return warns


def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(2)
    html, out = sys.argv[1], pathlib.Path(sys.argv[2])
    files = render(html, out)
    print(f"렌더 {len(files)}장 → {out}")
    print("컨택트 시트:", contact(files, out))
    if "--check" not in sys.argv:
        return
    if not CHECK_FILL.exists():
        print(f"[WARN] slide-density check_fill.py 없음({CHECK_FILL}) → R* 채움 검사 생략")
    fail = 0
    for f, lay, res in files:
        bad = res["bad"]
        for m in bad:
            print(f"[FAIL] {f.name} {m}")
        for m in contrast_warnings(f, res["texts"]):
            print(f"[WARN] {f.name} {m}")
        ok = not bad
        if lay.startswith("R") and CHECK_FILL.exists():
            r = subprocess.run([sys.executable, str(CHECK_FILL), str(f)], capture_output=True, text=True, encoding="utf-8", errors="replace")
            dens = r.returncode == 0
            if not dens:
                print(f"[FAIL] {f.name} 채움(slide-density)\n" + r.stdout.strip()[-600:])
            ok = ok and dens
        tag = "보고형 채움·DOM" if lay.startswith("R") else "키노트 DOM — contact.png 육안 검토 필요"
        print(f"[{'PASS' if ok else 'FAIL'}] {f.name} ({tag})")
        fail += 0 if ok else 1
    sys.exit(1 if fail else 0)


if __name__ == "__main__":
    main()
