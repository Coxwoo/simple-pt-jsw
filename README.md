# simple-pt — "Simple is best" 발표 자료 양식

- 짧은 영어 제목 + 결론 부제 (음슴체), 한 장에 그림 2–3 개, 빨간 결론 줄 하나, 근거는 Appendix.
- 디자인을 코드 (`tools/simple_pt.py`, python-pptx) 에 넣어서, Claude Code 든 ChatGPT 든 "이 양식으로 PT 만들어줘" 라고 하면 같은 모양의 pptx 가 나온다.
- 그래프는 함께 쓰는 저장소 [figstyle](https://github.com/Coxwoo/figstyle-jsw) 로 그린다.

![예시 덱 — 가상의 증폭기 "Demo Amp" (가짜 데이터)](example/preview/overview.png)

## 목차

1. 파일 — 무엇이 어디에
2. Claude Code 에서
3. ChatGPT 에서
4. 직접 — Python
5. 규칙 한눈에
6. 예시 덱
7. 글꼴
8. 라이선스

---

## 1. 파일

| 파일 | 내용 |
|---|---|
| `SKILL.md` | Claude Code skill — 만드는 순서와 꼭 지킬 것 |
| `CLAUDE.md` | 이 저장소를 Claude Code 로 열었을 때 |
| `chatgpt/INSTRUCTIONS.md` | ChatGPT Project 지시문과 올릴 파일 |
| `rules/design.md` | 모양 — 캔버스·글꼴·크기·색·레이아웃 |
| `rules/writing.md` | 글 — 제목·부제·빨간 줄·캡션, 음슴체, 숫자 |
| `rules/structure.md` | 구성 — 흐름, 장 수, Appendix, 만드는 순서 |
| `rules/checklist.md` | 검수 목록 |
| `tools/simple_pt.py` | pptx 를 만드는 모듈 — 디자인 값의 정본 |
| `tools/fix_pptx.py` | 다른 도구로 만든 pptx 의 후처리 — 한글 글꼴·줄바꿈·아래 첨자·노트 지우기 |
| `example/` | 예시 덱 — `make_example.py`, `demo_amp.pptx`, 그림, 미리보기 |

## 2. Claude Code 에서

- 모든 프로젝트에서 쓰기: 이 저장소와 figstyle 을 skill 폴더로 받는다. 그 뒤 "simple-pt 양식으로 PT 만들어줘".

  ```bash
  git clone https://github.com/Coxwoo/simple-pt-jsw.git ~/.claude/skills/simple-pt
  git clone https://github.com/Coxwoo/figstyle-jsw.git ~/.claude/skills/figstyle
  ```

- 이 저장소에서만: 폴더를 Claude Code 로 열면 `CLAUDE.md` 가 `SKILL.md` 를 가리킨다.
- 처음 한 번: `pip install python-pptx`.

## 3. ChatGPT 에서

- Project 를 만들고 `chatgpt/INSTRUCTIONS.md` 의 지시문을 Instructions 에 붙인다.
- 파일로 `tools/simple_pt.py`, `rules/` 네 개, `example/make_example.py` 를 올린다 (그래프도 그리면 [figstyle](https://github.com/Coxwoo/figstyle-jsw) 의 `figstyle.py`).
- ChatGPT 의 Python 에서 pptx 를 만들지 못하면, 스토리보드와 장 코드만 받아 내 PC 에서 돌린다.

## 4. 직접 — Python

```bash
pip install python-pptx
python example/make_example.py          # example/demo_amp.pptx
```

```python
import sys
sys.path.insert(0, "tools")
from simple_pt import Deck

d = Deck("My Talk")
d.cover("Short Topic:\nWhat It Does", subtitle="한 줄 부제",
        name="Your Name", affiliation="Dept. · Lab", date="YYYY-MM-DD")
d.panels("Results", "이전 설계 대비 이득 +14 dB",
         panels=[{"label": "Gain", "image": "figures/gain.png", "caption": "UGF 약 3 배"}],
         key="두 목표 모두 충족 (sim)", notes="말할 대본")
d.save("my_talk.pptx")                  # 규칙에 어긋난 곳은 경고로 알려 준다
```

## 5. 규칙 한눈에

| 글 | 규칙 |
|---|---|
| 제목 | 짧은 영어 주제 2–4 단어 |
| 부제 | 결론 한 문장, 음슴체·단정형 |
| 빨간 줄 | 굵은 결론 숫자 한 줄, 한 장에 하나 |
| 캡션 | 그림마다 무엇을 봐야 하는지 한 줄 |

- 흑백 미니멀, 강조는 빨강 하나. 본문은 짧게, 근거는 Appendix (한 장에 질문 하나) 와 노트로.
- 측정과 sim 을 섞지 않는다. 출처 줄과 "Next:" 연결 문장은 쓰지 않는다.
- 자세히는 `rules/`.

## 6. 예시 덱

- `example/demo_amp.pptx` — 가상의 증폭기 "Demo Amp", 가짜 데이터. 표지 → 본문 7 장 → Appendix 색인 → A1·A2.
- 장마다 뽑은 그림은 `example/preview/` (PowerPoint 로 뽑음).

## 7. 글꼴

- 영문 Arial, 한글 Noto Sans KR (Google Fonts 에서 무료). Noto Sans KR 이 없는 PC 에서는 한글이 맑은 고딕 등으로 바뀌어 보인다.
- 남에게 보낼 때 PDF 를 같이 내면 글꼴 걱정이 없다.

## 8. 라이선스

MIT — `LICENSE`.
