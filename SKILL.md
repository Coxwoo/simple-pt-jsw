---
name: simple-pt
description: 발표 자료 (PT·슬라이드·pptx) 를 "Simple is best" 양식으로 만든다 — 짧은 영어 제목 + 결론 부제 (음슴체), 한 장에 그림 2–3 개, 빨간 결론 줄 하나, 근거는 Appendix. "이 양식으로 PT 만들어줘", "발표 자료 만들어줘", "슬라이드 만들어줘", "덱 만들어줘" 같은 요청에 쓴다.
---

# simple-pt — "Simple is best" 발표 자료

이 skill 폴더 (Base directory) 에 규칙 `rules/`, pptx 를 만드는 `tools/simple_pt.py`, 후처리 `tools/fix_pptx.py`, 예시 `example/` 가 있다.

## 순서

1. 묻는다 — 청중, 발표 시간 (분), 핵심 주장 한 문장, 쓸 자료 (그림·숫자, 측정인지 sim 인지). 모르는 숫자는 지어내지 않는다.
2. `rules/structure.md`·`rules/writing.md` 를 읽고 스토리보드를 표로 낸다 — 장마다 제목 (영어 2–4 단어)·부제 (결론 한 문장)·그림·빨간 줄, 그리고 Appendix 질문 목록. 합의한 뒤에 만든다.
3. 그래프는 figstyle skill (같이 쓰는 저장소 https://github.com/Coxwoo/figstyle-jsw) 로 그린다. 새 그림은 후보와 추천을 먼저 내고 합의한 뒤에.
4. 장을 만든다 — `example/make_example.py` 를 본보기로:

   ```python
   import sys
   sys.path.insert(0, r"<이 skill 폴더>/tools")
   from simple_pt import Deck

   d = Deck("제목")
   d.cover(...)                       # 표지
   d.figure_bullets(...)              # 그림 + 불릿
   d.panels(...)                      # 그림 1–3 칸
   d.cards(...); d.table(...); d.summary(...)
   d.appendix_index([("질문", "A1")]) # 이 뒤의 장은 A1, A2 …
   d.save("deck.pptx")
   ```

   장 종류와 인자는 `tools/simple_pt.py` 머리말에 있다. 디자인 값 (글꼴·크기·색·자리) 을 장마다 바꾸지 않는다.
5. 검수 — save() 경고를 모두 고친다 (0 개). `rules/checklist.md` 를 확인한다. 가능하면 PowerPoint (또는 LibreOffice) 로 장을 그림으로 뽑아 겹침·잘림을 눈으로 본다.
6. 제출본이 필요하면 `python tools/fix_pptx.py deck.pptx --strip-notes` (노트를 지운 판).

## 꼭 지킬 것

- 제목은 짧은 영어 주제, 부제는 결론 한 문장 (최대 2 줄). 한글 문장은 음슴체·단정형 ("~사용", "~이내") — "~다." 로 끝내지 않는다.
- 한 장에 메시지 하나, 빨간 줄 하나, 그림 1–3 개. 본문은 늘리지 않고 근거는 Appendix·노트로.
- 측정과 sim 을 섞지 않고 sim 값에는 "(sim)". 출처 줄과 "Next:" 연결 문장은 쓰지 않는다.
- 노트에는 말할 대본 (존댓말 구어).
