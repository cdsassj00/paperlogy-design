# PPTX로 구현할 때 (python-pptx)

HTML 템플릿이 1순위다. 사용자가 .pptx 결과물을 요구할 때만 이 문서를 읽는다.
원본도 PowerPoint 제작물이므로 모든 기법이 네이티브로 재현된다.

## 기본 설정
```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
prs = Presentation(); prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)   # 16:9 먼저 확정
```
- 1280px 캔버스 좌표 → EMU: `Emu(int(px * 12192000 / 1280))`
- HTML px 글자크기 → pt: `pt = px * 0.75`

## 폰트 (Windows 레거시 패밀리명)
폰트 파일: `~/.claude/skills/paperlogy-design/assets/fonts/ttf/`. 받는 사람 PC에 없으면 대체 글꼴로 깨지므로 TTF를 함께 전달하거나 PDF도 같이 내보낸다.
본문은 Pretendard: `Pretendard Regular` / `Pretendard Medium` / `Pretendard SemiBold` / `Pretendard Bold`.
`Paperlogy 4 Regular` · `Paperlogy 5 Medium` · `Paperlogy 7 Bold` · `Paperlogy 8 ExtraBold` · `Paperlogy 9 Black` · 영문 `Montserrat Black`.
굵기는 폰트명으로 바꾼다(`run.font.bold=True` 금지). 한글은 `a:ea`, 영문은 `a:latin` 둘 다 지정:
```python
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
def set_font(run, name, size_pt, color=None, spc=None):
    run.font.size = Pt(size_pt); run.font.name = name          # a:latin
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))                                   # lxml 요소는 truthiness 금지 → None 비교
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
    ea.set('typeface', name)
    if spc is not None: rPr.set('spc', str(spc))          # 1/100pt. 40pt 제목 → -300, 20pt 본문 → -150
    if color: run.font.color.rgb = RGBColor.from_string(color)
```
- 제목 안 굵기 혼합 = run 분할: 키워드 run은 `Paperlogy 8 ExtraBold`, 조사 run은 `Paperlogy 4 Regular`.
- 마침표 run만 `FF2F92`.
- 행간 100%: `paragraph.line_spacing = 1.0` (본문 문단만 1.3–1.5).

## 배경
- 단색 검정: `slide.background.fill.solid(); ...fore_color.rgb = RGBColor(0,0,0)`
- 오로라/성운: 1) 원형(ellipse) 도형 + 방사형 그라데이션(중심 컬러 → 투명) + 소프트에지, 또는 2) PIL로 1920×1080 그라데이션 PNG를 생성해 배경 이미지로 깐다(가장 안정적). HTML 템플릿의 CSS 그라데이션 값을 그대로 쓴다.

## 시그니처 요소
| 요소 | 구현 |
|---|---|
| 태그 글로우 | 사각형 + `a:effectLst/a:outerShdw`(blurRad 228600, dist 76200, dir 5400000, 같은 컬러 alpha 60%) |
| 글래스 카드 | 둥근 사각형, 채우기 `#14161A` 투명도 25%, 선 `#FFFFFF` 투명도 92% 0.75pt |
| 못 박기 | 카드 네 모서리 안쪽 14px에 `+` 텍스트(Montserrat Bold 12pt) 또는 L자 선 2개 |
| 오브(구체) | 타원 + 경로형 그라데이션(#E7A6FF→#4B1A86→#000) |
| 후광 | 대형 타원 + 방사형 그라데이션(#965AFF α75% → 투명), 오브젝트 뒤로 |
| 사진 마스크 | 슬라이드 크기 사각형 + 선형 그라데이션(검정 α8% → α100% 투명), 사진 위 텍스트 아래 |
| 흑백 사진 | `a:blip`에 `<a:grayscl/>` 추가 |
| 순차 공개 | 애니메이션 대신 슬라이드 복제(페이지 쪼개기, P45) |

## 검증
PowerPoint COM 또는 LibreOffice(`soffice --headless --convert-to pdf`)로 PDF → PyMuPDF로 PNG → `contact.png` 육안 검토 + R(보고형) 장표는 `slide-density/scripts/check_fill.py`.
