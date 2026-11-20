# paperlogy-design

페이퍼로지 디자인 원칙(디자인 모드)과 원본 189장 편집 가능 PPTX 템플릿(템플릿 모드)을 하나로 묶은 에이전트 스킬입니다.

## 설치

**npx (Claude Code · Codex · Cursor 등 여러 에이전트 공통)**
```bash
npx skills add cdsassj00/paperlogy-design
```

**Claude Code 플러그인**
```
/plugin marketplace add cdsassj00/paperlogy-design
/plugin install paperlogy-design@paperlogy-design
```

**직접 복사**: `skills/paperlogy-design` 폴더를 `~/.claude/skills/`, `~/.agents/skills/`, `~/.codex/skills/` 중 쓰는 곳에 넣습니다.

## 구성
- `skills/paperlogy-design/SKILL.md`: 모드 선택 + 디자인 원칙 워크플로
- `assets/`: HTML 덱 틀(24 아키타입), Paperlogy·Pretendard 글꼴(OFL 1.1)
- `references/`: 61원칙, 레이아웃, 디자인 토큰, PPTX 가이드
- `scripts/`: 렌더·검사(render_slides, render_pptx, deck_lint, pptx_builder)
- `templates/`: 189장 템플릿 팩 5종(pptx), 카탈로그, use_template.py

## 주의
템플릿 팩 안의 사진·로고 등 제3자 자료는 재배포 권리가 별도입니다. 이 저장소는 비공개로 운영합니다.