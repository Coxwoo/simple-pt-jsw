"""equations — LaTeX 식을 슬라이드·논문용 그림으로 만든다 (matplotlib mathtext — 따로 설치할 것 없음).

    python tools/equations.py "A_{v} = g_{\\mathrm{m}} R_{\\mathrm{L}}" -o figures/eq/gain
    python tools/equations.py --file my_eq.txt -o figures/eq/my_eq        # 한 줄에 식 한 줄, 맞출 자리에 &

    from equations import render
    render([r"\\overline{v_{\\mathrm{n}}^{2}} &= \\frac{4kT\\gamma}{g_{\\mathrm{m}}}",
            r"g_{\\mathrm{m}} &= \\frac{2I_{\\mathrm{D}}}{V_{\\mathrm{ov}}}"], "figures/eq/noise")

- 글꼴은 Computer Modern — LaTeX·IEEE 논문과 같은 모양. 크기 16 pt (슬라이드 본문과 어울린다).
- 분수는 LaTeX 의 display 식처럼 큰 꼴로 그린다 (\\frac → \\dfrac).
- 여러 줄은 줄마다 쓰고, 맞출 자리 (보통 = 앞) 에 `&` 를 둔다 — LaTeX 의 aligned 처럼.
- 쓰는 것:
    <이름>.png  투명 배경, 600 dpi — simple_pt 의 equation() 이 이 크기 그대로 (16 pt 상당) 넣는다
    <이름>.svg  글자는 윤곽선 — 그림 편집·논문용
    <이름>.tex  원본 — 여러 줄이면 aligned 환경. 논문에 \\input 으로 쓴다
- matplotlib mathtext 가 아는 LaTeX 만 된다 (첨자, \\frac, \\sqrt, \\sum, \\int, 그리스 문자, \\mathrm,
  \\left( \\right), \\overline, \\approx, \\, …). \\text{} 대신 \\mathrm{}. 표기 규칙은 rules/equations.md.
"""
import argparse
import re
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

SIZE_PT, DPI, INK = 16, 600, "#111111"
LINE_GAP = 0.55                      # 줄 사이 빈 높이 — 글자 크기의 배수
RC = {"mathtext.fontset": "cm", "svg.fonttype": "path", "svg.hashsalt": "equations"}


def _split(latex):
    """식을 줄 목록으로 — 리스트, 또는 줄바꿈이나 LaTeX 의 \\\\ 로 나눈 글."""
    if isinstance(latex, str):
        latex = latex.replace("\\\\", "\n").splitlines()
    return [ln.strip() for ln in latex if ln.strip()]


def render(latex, out, size=SIZE_PT, color=INK):
    """식을 <out>.png·.svg·.tex 로 쓴다. 돌려주는 것: PNG 경로. 읽지 못하는 식이면 ValueError (어느 줄인지 알려 준다)."""
    lines = _split(latex)
    if not lines:
        raise ValueError("식이 비었다")
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    aligned = any("&" in ln for ln in lines)
    with plt.rc_context(RC):
        fw, fh = 14, 2 + 1.2 * len(lines)                   # 넉넉한 그림 (in) — 저장할 때 식 둘레로 자른다
        fig = plt.figure(figsize=(fw, fh), dpi=DPI)
        r = fig.canvas.get_renderer()
        x0 = fw / 2                                         # 맞추는 자리 (in) — 그림 가운데
        rows = []
        for ln in lines:
            if aligned and "&" in ln:
                left, right = ln.split("&", 1)
                parts = [(left, "right"), (right, "left")]
            else:
                parts = [(ln, "left" if aligned else "center")]
            row = []
            for tex, ha in parts:
                if tex.strip():
                    # 분수는 큰 꼴 (\dfrac) — LaTeX 의 display 식과 같게. .tex 원본은 \frac 그대로 둔다
                    shown = re.sub(r"\\frac(?![A-Za-z])", r"\\dfrac", tex.strip())
                    # 자리는 그림 비율 (transFigure) 로 준다 — 저장할 때 자르는 것 (bbox_inches) 이 이 좌표만 옮긴다
                    t = fig.text(x0 / fw, 0, f"${shown}$", fontsize=size, color=color, ha=ha, va="baseline")
                    try:
                        t.get_window_extent(r)
                    except ValueError as e:
                        plt.close(fig)
                        why = [s for s in str(e).splitlines() if s.strip()]
                        raise ValueError(f"식을 읽지 못함: {ln}\n  {why[-1] if why else ''}") from None
                    row.append(t)
            rows.append(row)
        # 줄마다 기준선 위 (ascent)·아래 (descent) 를 재고, 위에서부터 차례로 놓는다
        y = fig.get_figheight() - 1.0
        gap = size * LINE_GAP / 72
        prev_descent = 0.0
        for i, row in enumerate(rows):
            ext = [t.get_window_extent(r) for t in row]          # 지금 기준선은 0 in 에 있다
            ascent = max(e.y1 for e in ext) / DPI
            descent = -min(e.y0 for e in ext) / DPI
            if i:
                y -= prev_descent + gap + ascent
            for t in row:
                t.set_position((x0 / fw, y / fh))
            prev_descent = descent
        fig.savefig(out.with_suffix(".png"), dpi=DPI, transparent=True, bbox_inches="tight", pad_inches=0.02,
                    metadata={"Software": None})
        fig.savefig(out.with_suffix(".svg"), transparent=True, bbox_inches="tight", pad_inches=0.02,
                    metadata={"Date": None, "Creator": None})
        plt.close(fig)
    env = "aligned" if aligned else "gathered"
    body = lines[0] if len(lines) == 1 else f"\\begin{{{env}}}\n" + " \\\\\n".join(lines) + f"\n\\end{{{env}}}"
    out.with_suffix(".tex").write_text(body + "\n", encoding="utf-8")
    return out.with_suffix(".png")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="LaTeX 식 → PNG (슬라이드)·SVG·TeX")
    parser.add_argument("latex", nargs="?", help="식 (한 줄). 여러 줄은 \\\\ 로 나누고 맞출 자리에 &")
    parser.add_argument("--file", type=Path, help="식 파일 — 한 줄에 식 한 줄")
    parser.add_argument("-o", "--out", required=True, help="확장자 없는 출력 경로 (예: figures/eq/gain)")
    parser.add_argument("--size", type=float, default=SIZE_PT, help=f"글자 크기 pt (기본 {SIZE_PT})")
    args = parser.parse_args()
    if not args.latex and not args.file:
        parser.error("식이나 --file 을 준다")
    latex = args.file.read_text(encoding="utf-8").splitlines() if args.file else args.latex
    try:
        png = render(latex, args.out, args.size)
    except ValueError as e:
        print(e)
        sys.exit(1)
    from PIL import Image
    with Image.open(png) as im:
        w, h = im.size
    print(f"{png} — {w / DPI * 72:.0f} × {h / DPI * 72:.0f} pt (글자 {args.size:g} pt), .svg·.tex 도 썼다")


if __name__ == "__main__":
    main()
