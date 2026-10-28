# 식 — 쓰는 법과 넣는 법

- 식은 LaTeX 로 쓰고 `tools/equations.py` 로 그림을 만든다. 엔진은 MathJax 3 (mathjax-full 3.2.2 으로 고정) — 글자를 path 로 그려 글꼴이 필요 없고, 어느 PC 에서나 같은 그림이 나온다.
- 표기는 IEEE 논문 관례를 따른다 — 변수는 이탤릭, 뜻을 나타내는 첨자와 단위는 바로 선 글자.
- 슬라이드에는 `d.equation()` 으로 넣는다 — 회색 수식 박스 가운데, 번호 없음, 바로 아래 where 한 문장.

## 목차

1. 어디에 쓰나 — 본문과 Appendix
2. IEEE 표기 — 표
3. 만들기 — tools/equations.py (MathJax)
4. 슬라이드에 넣기 — d.equation()
5. 검산 — SymPy (선택)
6. 준비와 막힐 때

---

## 1. 어디에 쓰나

- 본문에는 확실한 식만 쓴다 (정의, 교과서 식, 측정으로 확인한 관계). 가정이 큰 모형 식과 유도는 Appendix 로.
- 한 장에 식 박스 하나. 여러 줄이면 = 자리에 맞춘다.
- 출처 표기는 쓰지 않는다.

## 2. IEEE 표기

| 무엇 | 규칙 | 예 (LaTeX) |
|---|---|---|
| 변수 | 이탤릭 (기본) | `V`, `g`, `C` |
| 뜻을 나타내는 첨자 | 바로 선 글자 | `V_{\mathrm{IN}}`, `g_{\mathrm{m}}`, `C_{\mathrm{L}}` |
| 번호·변수 첨자 | 이탤릭 | `v_{i}`, `x_{k}` |
| 단위·함수·글 | 바로 선 글자, 숫자와 단위 사이 얇은 빈칸 `\,` | `10\,\mathrm{pF}`, `\log`, `\text{when}` |
| 곱 | 붙여 쓰거나 `\cdot` | `g_{\mathrm{m}} R_{\mathrm{L}}` |
| 분수 | `\frac` (display 꼴로 그려진다) | `\frac{4kT}{g_{\mathrm{m}}}` |
| 괄호 크기 | `\left( … \right)` | `\left( \frac{a}{b} \right)^{2}` |
| 여러 줄 | `aligned` 로 = 에 맞춘다 (줄마다 쓰고 `&` 만 두면 알아서 묶는다) | `A_{v} &= g_{\mathrm{m}} R_{\mathrm{L}}` |

- 한 줄이 IEEE 두 단 논문의 한 칸 (252 pt) 보다 넓으면 `equations.py` 가 알린다 — 논문에 쓸 식이면 줄을 나눈다.

## 3. 만들기

```bash
python tools/equations.py "A_{v} = g_{\mathrm{m}} R_{\mathrm{L}}" -o figures/eq/gain
python tools/equations.py --file figures/eq/noise.tex -o figures/eq/noise      # 식 본문 파일 (% 줄은 주석)
```

여러 줄은 Python 에서 리스트로 (또는 `\begin{aligned} … \end{aligned}` 를 그대로):

```python
import sys
sys.path.insert(0, "tools")
from equations import render

render([r"\overline{v_{\mathrm{n}}^{2}} &= \frac{4kT\gamma}{g_{\mathrm{m}}}",
        r"g_{\mathrm{m}} &= \frac{2I_{\mathrm{D}}}{V_{\mathrm{ov}}}"], "figures/eq/noise")
```

| 결과 | 크기 | 쓰는 곳 |
|---|---|---|
| `noise.svg` | IEEE 본문 10 pt 기준 | 논문·그림 편집 |
| `noise.png` | 1.6 배 (16 pt 상당), 600 dpi, 투명 배경 | 슬라이드 (`d.equation()`) |
| `noise.tex` | 원본 | 논문에 `\begin{equation}\input{noise.tex}\end{equation}` |

- 핵심 식 하나를 더 크게 보이려면 `--scale 2` (20 pt 상당).
- 파일 이름은 내용으로 (`noise`, `gain`). 장 번호는 넣지 않는다.
- 만든 PNG 를 열어 본다 — 첨자가 바로 섰는지, = 자리가 맞는지.

## 4. 슬라이드에 넣기

```python
d.equation("Key Idea", "부제 — 결론 한 문장",
           eq="figures/eq/noise.png",                      # LaTeX 글이나 줄 리스트를 바로 줘도 된다
           where="k 볼츠만 상수, T 절대 온도, γ 잡음 계수",
           cards=[{"title": "…", "body": ["…"]}], key="빨간 결론 줄")
```

- 식은 회색 박스 가운데에 원래 크기 (16 pt 상당) 로 들어간다. 번호는 달지 않는다.
- where 는 한 문장, 2 줄 이하. 기호 설명을 쉼표로 이어 쓴다.
- 박스 아래에는 그림 하나 (왼쪽) 와 카드 2–3 개 (오른쪽), 또는 불릿을 둔다. 그림이 없으면 카드가 전체 폭을 쓴다.
- 글 속의 짧은 식은 첨자 표시로 충분하다 — `V_{IN}`, `x^{2}`.

## 5. 검산 — SymPy (선택)

- 유도한 식은 숫자를 넣어 보거나 SymPy 로 양쪽이 같은지 확인한다 (`pip install sympy`).

```python
import sympy as sp

R1, R2 = sp.symbols("R1 R2", positive=True)
assert sp.simplify(1 / (1 / R1 + 1 / R2) - R1 * R2 / (R1 + R2)) == 0   # 병렬 저항 식
```

## 6. 준비와 막힐 때

- 처음 한 번: Node.js (LTS) 를 설치하고 `python check_setup.py --install` — `tools/mathjax` 에서 `npm ci` 로 MathJax 를 받는다. PNG 는 PyMuPDF (`pip install pymupdf`) 가 만든다.
- MathJax 는 LaTeX 수식 문법을 거의 다 안다 — `aligned`, `\text`, `\operatorname`, `\newcommand` 까지. 문서 명령 (`\section` 등) 과 그림은 안 된다.
- 식이 틀리면 `equations.py` 가 MathJax 의 이유를 그대로 알려 준다 (예: `Undefined control sequence \unknowncmd`).
- `npm ci` 때 나오는 `deprecated` 경고는 판을 3.2.2 로 고정해서 나는 것이다 — 모양을 그대로 지키려고 바꾸지 않는다.
