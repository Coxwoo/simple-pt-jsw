# AGENTS.md — Codex 용 안내

이 저장소는 Claude Code 를 기준으로 쓴 발표 자료 양식이다. Codex 도 같은 파일을 정본으로 삼아 같은 결과를 낸다.

- 할 일의 순서는 `SKILL.md` 다 — 준비 (설치 확인) → 묻기 → 스토리보드 합의 → 그림·식 → 장 만들기 → 검수 → 제출본.
- 규칙은 `rules/` (design·writing·structure·equations·checklist). 디자인 값은 `tools/simple_pt.py` 한 곳에 있으니 바꾸지 않는다.
- 처음에는 `python check_setup.py` 로 설치를 확인한다. 빠진 Python 패키지는 `--install`, 글꼴은 안내를 보여 주고 사용자 확인 뒤 설치를 돕는다.
- 그래프는 figstyle 저장소 (https://github.com/Coxwoo/figstyle-jsw) 로 그린다.
- Claude Code 에만 있는 것 — skill 폴더 (`~/.claude/skills/…`), `CLAUDE.md`, "이 skill 폴더 (Base directory)" 같은 말, Claude 의 도구 이름 — 은 Codex 환경에 맞게 스스로 판단해 같은 일을 한다 (예: skill 폴더 = 이 저장소 폴더).
- 판단이 갈리는 곳 (설치 권한, 네트워크, 장을 그림으로 확인하는 방법) 은 사용자에게 알리고 정한다.
