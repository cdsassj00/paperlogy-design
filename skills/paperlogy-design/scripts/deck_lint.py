# -*- coding: utf-8 -*-
"""
deck_lint.py — 덱의 '구성·리듬·문장량'을 검사한다 (render_pptx.py가 못 잡는 것).

    python deck_lint.py deck.pptx              # 강의·발표(KEYNOTE 혼합) 기준
    python deck_lint.py deck.pptx --report     # 사내 보고형(REPORT 위주) 기준
    python deck_lint.py deck.pptx --outline outline.md   # 원고 파일도 함께 검사

render_pptx.py --check 는 '넘침 · 겹침 · 글자 크기 · 빈 띠'를 본다. 그건 PASS여도
'같은 장표 복붙 · 문단 붙여넣기 · 기승전결 없음'은 그대로 통과한다. 이 스크립트가 그걸 잡는다.

FAIL (고칠 때까지 납품 금지)
  L1 한 레이아웃이 전체의 35% 초과                (강의) / 60% 초과 (--report)
  L2 같은 레이아웃 4장 이상 연속
  L3 레이아웃 종류 부족: 20장 이상인데 8종 미만
  L4 보고형(R*) 비율 50% 초과 (강의 모드)
  T1 한 장의 화면 글자 수 > 520자 (발표자 메모로 옮길 내용)
  T2 한 텍스트 상자 > 220자 (= 문단 붙여넣기)
  T3 덱 평균 화면 글자 수 > 230자 (강의) / 380자 (--report)
  N1 강의 모드인데 발표자 메모(≥60자)가 있는 장이 60% 미만
  O1 --outline: 원고 줄 수 ≠ 장 수, 또는 장당 메시지가 40자 초과
WARN
  T4 제목(가장 큰 글자 상자) 35자 초과
  L5 다크(검정) 배경만 4장 이상 연속 — 명암 리듬
레이아웃 ID는 deck.pptx 옆 deck.layouts.json(pptx_builder가 저장)에서 읽는다.
"""
import sys, os, json, re, argparse, statistics
from pptx import Presentation
from pptx.util import Emu

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ap = argparse.ArgumentParser()
ap.add_argument("pptx")
ap.add_argument("--report", action="store_true", help="사내 보고형 덱 기준(완화)")
ap.add_argument("--outline", help="장당 한 문장 원고 파일(md/txt). '- '·'NN.' 줄을 메시지로 센다")
a = ap.parse_args()

prs = Presentation(a.pptx)
slides = list(prs.slides)
n = len(slides)
lj = os.path.splitext(a.pptx)[0] + ".layouts.json"
layouts = json.load(open(lj, encoding="utf-8")) if os.path.exists(lj) else ["?"] * n
H = prs.slide_height
fails, warns = [], []


def F(code, msg):
    fails.append(f"[FAIL] {code} {msg}")


def W(code, msg):
    warns.append(f"[WARN] {code} {msg}")


def is_chrome(shape):
    """헤더·푸터·쪽번호처럼 화면 가장자리의 작은 글자는 제외"""
    if shape.top is None:
        return False
    top, bottom = shape.top, shape.top + shape.height
    size = 0
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                size = max(size, r.font.size.pt)
    small = size and size <= 9.5
    return small and (bottom < H * 0.09 or top > H * 0.93)


per_slide, dark = [], []
for i, s in enumerate(slides):
    boxes = []
    for sh in s.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t and not is_chrome(sh):
                size = max([r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size] or [0])
                boxes.append((t, size))
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    if c.text.strip():
                        boxes.append((c.text.strip(), 10))
    chars = sum(len(re.sub(r"\s", "", t)) for t, _ in boxes)
    longest = max(boxes, key=lambda b: len(b[0]), default=("", 0))
    title = max(boxes, key=lambda b: b[1], default=("", 0))[0]
    note = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
    bg = s.background.fill
    try:
        is_dark = bg.type is not None and str(bg.fore_color.rgb) in ("000000",)
    except Exception:
        is_dark = False
    # 붙여 넣으라고 보여 주는 프롬프트 카드(‘PROMPT’ 라벨)는 길어도 정상 → T2 면제
    if any(t.strip().upper().startswith("PROMPT") for t, _ in boxes):
        longest = ("", 0)
    per_slide.append(dict(i=i + 1, lay=layouts[i] if i < len(layouts) else "?", chars=chars,
                          longest=len(re.sub(r"\s", "", longest[0])), longtxt=longest[0][:40],
                          title=title.replace("\n", " "), note=len(note)))
    dark.append(is_dark)

# ── 레이아웃 다양성 · 리듬 ──
from collections import Counter
cnt = Counter(layouts[:n])
top_lay, top_n = cnt.most_common(1)[0]
lim = 0.60 if a.report else 0.35
if top_n / n > lim:
    F("L1", f"'{top_lay}' 레이아웃이 {top_n}/{n}장({top_n / n:.0%}) — 한도 {lim:.0%}. 같은 틀에 내용만 바꿔 넣은 덱입니다.")
run, prev, worst = 0, None, (0, None, 0)
for i, l in enumerate(layouts[:n]):
    run = run + 1 if l == prev else 1
    prev = l
    if run > worst[0]:
        worst = (run, l, i + 1)
if worst[0] >= 4:
    F("L2", f"'{worst[1]}' {worst[0]}장 연속 (~{worst[2]}장). 3장 이하로 끊고 키노트 장표(K02/K06/K10…)를 섞으세요.")
if n >= 20 and len(cnt) < 8:
    F("L3", f"레이아웃 {len(cnt)}종뿐 ({n}장). 8종 이상 쓰세요.")
r_ratio = sum(v for k, v in cnt.items() if str(k).startswith("R")) / n
if not a.report and r_ratio > 0.5:
    F("L4", f"보고형(R*) {r_ratio:.0%} — 강의 덱은 50% 이하. 근거·실습만 R, 메시지·전환은 K.")
drun = 0
for i, dk in enumerate(dark):
    drun = drun + 1 if dk else 0
    if drun == 4:
        W("L5", f"검정 배경 4장 연속 (~{i + 1}장) — paper/holo 장표로 명암을 끊으세요.")

# ── 문장량 ──
for p in per_slide:
    if p["chars"] > 520:
        F("T1", f"{p['i']:02d}장 화면 글자 {p['chars']}자 — 설명은 발표자 메모로. ({p['title'][:24]})")
    if p["longest"] > 220:
        F("T2", f"{p['i']:02d}장 한 상자에 {p['longest']}자 — 문단 붙여넣기. 한 문장으로 줄이거나 쪼개세요. (\"{p['longtxt']}…\")")
    if len(p["title"]) > 35:
        W("T4", f"{p['i']:02d}장 제목 {len(p['title'])}자 — 30자 안팎으로. ({p['title'][:30]}…)")
avg = statistics.mean(p["chars"] for p in per_slide)
alim = 380 if a.report else 230
if avg > alim:
    F("T3", f"평균 화면 글자 {avg:.0f}자/장 — 한도 {alim}자. 원고를 장표에 옮기지 말고 '장당 한 문장'으로.")
if not a.report and n >= 10:
    with_notes = sum(1 for p in per_slide if p["note"] >= 60)
    if with_notes / n < 0.6:
        F("N1", f"발표자 메모 있는 장 {with_notes}/{n} — 강의 덱은 화면=한 문장, 말할 내용=메모.")

# ── 원고 ──
if a.outline:
    lines = [l.strip() for l in open(a.outline, encoding="utf-8") if re.match(r"^\s*(-|\d{1,2}[.)])\s+\S", l)]
    if len(lines) != n:
        F("O1", f"원고 메시지 {len(lines)}줄 ≠ 장 수 {n}.")
    for k, l in enumerate(lines, 1):
        msg = re.sub(r"^\s*(-|\d{1,2}[.)])\s+", "", l).split("|")[0].strip()
        if len(msg) > 40:
            F("O1", f"원고 {k:02d}: 메시지 {len(msg)}자 — 한 문장 40자 이내. ({msg[:30]}…)")

# ── 출력 ──
print(f"덱: {os.path.basename(a.pptx)} · {n}장 · 모드: {'보고형' if a.report else '강의/발표'}")
print(f"레이아웃 {len(cnt)}종 · 최다 '{top_lay}' {top_n}장({top_n / n:.0%}) · 보고형 {r_ratio:.0%} · 최장 연속 {worst[0]}장")
print(f"화면 글자 평균 {avg:.0f}자/장 · 최대 {max(p['chars'] for p in per_slide)}자 · 메모 있는 장 {sum(1 for p in per_slide if p['note'] >= 60)}/{n}")
for m in fails + warns:
    print(m)
print("\n결과:", "FAIL ❌ — 위 항목을 고친 뒤 다시 실행" if fails else "PASS ✅ (WARN은 육안 판단)")
sys.exit(1 if fails else 0)
