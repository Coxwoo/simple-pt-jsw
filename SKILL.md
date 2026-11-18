---
name: simple-pt
description: 발표 자료 (PT·슬라이드·pptx) 를 "Simple is best" 양식으로 만든다 — 짧은 영어 제목 + 질문·평서문을 섞은 부제 (음슴체), 그림 중심 (카드·회색 상자 없음, 논문식 표), 빨간 결론 줄은 파트를 맺는 장에만, IEEE 표기 식, 근거는 Appendix. "이 양식으로 PT 만들어줘", "발표 자료 만들어줘", "슬라이드 만들어줘", "덱 만들어줘", "슬라이드에 식 넣어줘" 같은 요청에 쓴다.
---

# simple-pt — "Simple is best" 발표 자료

이 skill 폴더 (Base directory) 에 규칙 `rules/`, pptx 를 만드는 `tools/simple_pt.py`, 식 그림을 만드는 `tools/equations.py`, 후처리 `tools/fix_pptx.py`, 설치 확인 `check_setup.py`, 예시 `example/` 가 있다.

## 0. 준비 — 이 PC 에서 처음이면

1. `python <이 skill 폴더>/check_setup.py` 를 돌린다.
2. 빠진 Python 패키지·MathJax 가 있으면 사용자에게 한 줄 알리고 `python <이 skill 폴더>/check_setup.py --install` (pip 설치와 `tools/mathjax` 의 `npm ci`).
3. 빠진 Node.js·글꼴 (Noto Sans KR 등) 은 스크립트가 알려 주는 공식 배포처 안내를 사용자에게 보여 주고, 사용자 확인 뒤 설치를 돕는다. 설치한 뒤 `--install` 로 마저 하고 (MathJax, matplotlib 글꼴 목록) 다시 확인한다.
4. figstyle skill 이 없으면 받게 안내한다: `git clone https://github.com/Coxwoo/figstyle-jsw.git ~/.claude/skills/figstyle`.
5. 다 갖추지 못해도 만들 수는 있다 — 무엇이 달라지는지 (글꼴이 바뀌어 모양·줄바꿈이 달라짐) 사용자에게 알린다.

## 1. 만드는 순서

1. 묻는다 — 청중, 발표 시간 (분), 핵심 주장 한 문장, 쓸 자료 (그림·숫자, 측정인지 sim 인지). 모르는 숫자는 지어내지 않는다.
2. `rules/structure.md`·`rules/writing.md` 를 읽고 스토리보드를 표로 낸다 — 장마다 제목 (영어 2–4 단어)·부제 (질문이나 결론 — 덱 전체에서 섞어서)·그림·꼴 (같은 꼴이 이어지지 않게)·빨간 줄 (파트를 맺는 장에만), 그리고 Appendix 질문 목록. 합의한 뒤에 만든다.
3. 그래프는 figstyle skill 로 그린다 (한 줄에 셋처럼 작게 넣어 글자가 너무 작아지면 큰 글자 판 `fs.big()`). 새 그림은 후보와 추천을 먼저 내고 합의한 뒤에.
4. 식은 `rules/equations.md` 대로 — IEEE 표기로 LaTeX 를 쓰고 `tools/equations.py` (MathJax) 로 그림을 만든다 (또는 `d.equation()` 에 LaTeX 를 바로 준다).
5. 장을 만든다 — `example/make_example.py` 를 본보기로:

   ```python
   import sys
   sys.path.insert(0, r"<이 skill 폴더>/tools")
   from simple_pt import Deck

   d = Deck("제목", parts=["Why", "Design", "Result"])     # parts — 오른쪽 위 진행 표시 (선택)
   d.cover(...)                                              # 표지
   d.outline([("Why", "한 줄"), ...])                        # 목차 (선택)
   d.divider("Why")                                          # 파트 구분 장 — 가운데 제목 + 밑줄 (선택)
   d.figure_bullets(..., part="Why", takeaway="굵은 결론 한 줄", caption="회색 조건")
   d.equation(..., eq=[r"A_{v} &= g_{\mathrm{m}} R_{\mathrm{L}}"], where="…")      # 채움 없는 검정 상자
   d.panels(..., key="빨간 줄 — 파트를 맺는 장에만")          # 그림 1–3 칸
   d.table(..., highlight=[(0, 1)], refs=["[1] …"])          # 논문식 표 + 참고문헌 줄
   d.appendix_index([("질문", "A1")])                        # 이 뒤의 장은 A1, A2 …
   d.save("deck.pptx")
   ```

   장 종류와 인자는 `tools/simple_pt.py` 머리말에 있다. 디자인 값 (글꼴·크기·색·자리) 을 장마다 바꾸지 않는다.
6. 검수 — save() 경고를 모두 고친다 (0 개). `rules/checklist.md` 를 확인한다. 장을 그림으로 뽑아 겹침·잘림을 눈으로 본다 (Windows 는 PowerPoint COM, 그 밖에는 LibreOffice `soffice --headless --convert-to pdf`).
7. 제출본이 필요하면 `python <이 skill 폴더>/tools/fix_pptx.py deck.pptx --strip-notes` (노트를 지운 판).

## 꼭 지킬 것

- 제목은 짧은 영어 주제, 부제는 그 장의 질문이나 결론 한 문장 (최대 2 줄) — 덱 전체에서 질문과 평서문을 적절히 섞는다. 한글 문장은 음슴체·단정형 ("~사용", "~이내") — "~다." 로 끝내지 않는다. 영어 덱이면 마침표 없는 짧은 구.
- 한 장에 메시지 하나, 그림 1–3 개. 빨간 줄은 파트를 맺는 장에만 — 다른 장의 결론은 근거 옆 (그림 아래 굵은 한 줄). 본문은 늘리지 않고 근거는 Appendix·노트로.
- AI 가 만든 것처럼 보이는 꼴을 쓰지 않는다 — 카드·회색 상자·큰 숫자 타일·단계 번호, 모든 장에 같은 틀, 그림 없는 내용 장. 상자가 꼭 필요하면 채움 없는 검정 2¼ pt, 표는 논문식.
- 그림마다 무엇을 어떤 조건으로 잰 것인지 슬라이드에서 보이게 (이름표·캡션).
- 식은 IEEE 표기 (변수 이탤릭, 뜻을 나타내는 첨자·단위는 바로 선 글자), 본문에는 확실한 식만, 번호 없음, 아래 where 한 문장.
- 측정과 sim 을 섞지 않고 sim 값에는 "(sim)". 인용은 [n] + 아래 작은 글씨 참고문헌 — 출처 줄과 "Next:" 연결 문장은 쓰지 않는다.
- 노트: 본문 장은 말할 대본만 (언어와 관계없이), Appendix 장은 질문 받을 때 말할 거리를 짧게.
- 사람이 고친 판이 있으면 그 위에 짓고 덮어쓰지 않는다. 고치라고 하지 않은 장은 고치지 않는다.
