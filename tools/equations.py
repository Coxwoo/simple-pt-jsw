"""equations — LaTeX 식을 슬라이드·논문용 그림으로 만든다 (MathJax 3).

    python tools/equations.py "A_{v} = g_{\\mathrm{m}} R_{\\mathrm{L}}" -o figures/eq/gain
    python tools/equations.py --file my_eq.tex -o figures/eq/my_eq          # 식 본문 파일 (% 로 시작하는 줄은 주석)

    from equations import render
    render([r"\\overline{v_{\\mathrm{n}}^{2}} &= \\frac{4kT\\gamma}{g_{\\mathrm{m}}}",
            r"g_{\\mathrm{m}} &= \\frac{2I_{\\mathrm{D}}}{V_{\\mathrm{ov}}}"], "figures/eq/noise")

- 엔진: MathJax 3 (mathjax-full 3.2.2, tools/mathjax). 처음 한 번 `python check_setup.py --install` (또는 tools/mathjax 에서
  `npm ci`). Node.js 가 필요하다. 글자는 path 라 글꼴이 필요 없고, 판을 package-lock.json 으로 고정해 어느 PC 에서나 같은 그림이 나온다.
- 여러 줄은 줄마다 쓰고 맞출 자리 (보통 = 앞) 에 `&` — aligned 로 묶는다 (& 가 없으면 gathered). \\begin{…} 이 든 글은 그대로 쓴다.
- 쓰는 것:
    <이름>.svg  IEEE 본문 10 pt 기준 크기 — 논문·그림 편집용. 폭이 IEEE 두 단의 한 칸 (252 pt) 을 넘으면 알린다
    <이름>.png  슬라이드용 — 1.6 배 (16 pt 상당), 600 dpi, 투명 배경. simple_pt 의 equation() 이 이 크기 그대로 넣는다
    <이름>.tex  원본 — 논문에 \\begin{equation}\\input{<이름>.tex}\\end{equation}
- PNG 는 PyMuPDF 가 SVG 를 그려 만든다 (pip install pymupdf). 표기 규칙은 rules/equations.md.
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

MATHJAX = Path(__file__).resolve().parent / "mathjax"
BODY_PT = 10.0                 # IEEE 본문 글자 크기
COLUMN_PT = 252.0              # IEEE 2-column 한 칸 (3.5 in)
SLIDE_SCALE, DPI = 1.6, 600    # 슬라이드 PNG — 본문의 1.6 배 (16 pt 상당), 600 dpi


def _tex(latex):
    """식 → TeX 본문. 리스트·줄바꿈·\\\\ 로 줄을 나누고, & 가 있으면 aligned, 없으면 gathered 로 묶는다."""
    if isinstance(latex, str) and "\\begin{" in latex:
        return latex.strip()
    lines = latex.replace("\\\\", "\n").splitlines() if isinstance(latex, str) else [str(x) for x in latex]
    lines = [ln.strip() for ln in lines if ln.strip() and not ln.lstrip().startswith("%")]
    if not lines:
        raise ValueError("식이 비었다")
    if len(lines) == 1:
        return lines[0]
    env = "aligned" if any("&" in ln for ln in lines) else "gathered"
    return f"\\begin{{{env}}}\n" + " \\\\\n".join(lines) + f"\n\\end{{{env}}}"


def _mathjax(tex):
    """TeX → MathJax SVG (tools/mathjax/render.cjs)."""
    node = shutil.which("node")
    if not node:
        raise RuntimeError("Node.js 가 없다 — python check_setup.py 의 안내대로 설치한다")
    if not (MATHJAX / "node_modules" / "mathjax-full").is_dir():
        raise RuntimeError("MathJax 가 설치되지 않았다 — python check_setup.py --install (또는 tools/mathjax 에서 npm ci)")
    r = subprocess.run([node, str(MATHJAX / "render.cjs")], input=json.dumps([{"name": "eq", "tex": tex}]),
                       capture_output=True, text=True, encoding="utf-8", cwd=MATHJAX)
    if r.returncode:
        raise RuntimeError(f"MathJax 실패\n{r.stderr}")
    out = json.loads(r.stdout)[0]
    if "error" in out:
        raise ValueError(f"식을 읽지 못함: {out['error']}\n  {tex}")
    return out["svg"]


def _standalone(svg):
    """MathJax SVG 의 크기를 pt 로 (본문 10 pt 기준), 색은 검정. → (svg, 폭 pt, 높이 pt)."""
    vb = [float(v) for v in re.search(r'viewBox="([^"]+)"', svg).group(1).split()]
    w, h = vb[2] / 1000 * BODY_PT, vb[3] / 1000 * BODY_PT          # viewBox 단위는 1/1000 em
    svg = re.sub(r' style="[^"]*"', "", svg, count=1)
    svg = re.sub(r'width="[^"]+"', f'width="{w:.2f}pt"', svg, count=1)
    svg = re.sub(r'height="[^"]+"', f'height="{h:.2f}pt"', svg, count=1)
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + svg.replace("currentColor", "#000000") + "\n", w, h


def render(latex, out, scale=SLIDE_SCALE):
    """식을 <out>.svg·.png·.tex 로 쓴다. 돌려주는 것: PNG 경로. 읽지 못하는 식이면 ValueError (MathJax 의 이유와 함께)."""
    try:
        import pymupdf
    except ImportError:                                       # 옛 판 PyMuPDF 의 이름
        import fitz as pymupdf
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tex = _tex(latex)
    svg, w, h = _standalone(_mathjax(tex))
    out.with_suffix(".svg").write_text(svg, encoding="utf-8")
    out.with_suffix(".tex").write_text(tex + "\n", encoding="utf-8")
    doc = pymupdf.open(out.with_suffix(".svg"))
    zoom = scale * DPI / 72                                   # SVG 는 pt 단위, 1 pt = 1/72 in
    pix = doc[0].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=True)
    pix.set_dpi(DPI, DPI)
    pix.save(out.with_suffix(".png"))
    doc.close()
    if w > COLUMN_PT:
        print(f"  {out.name}: 폭 {w:.0f} pt — IEEE 한 칸 (252 pt) 보다 넓다. 논문에 쓰면 줄을 나눈다")
    return out.with_suffix(".png")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="LaTeX 식 → SVG (논문, IEEE 10 pt)·PNG (슬라이드, 16 pt 상당)·TeX — MathJax 3")
    parser.add_argument("latex", nargs="?", help="식. 여러 줄은 \\\\ 로 나누고 맞출 자리에 &")
    parser.add_argument("--file", type=Path, help="식 본문 파일 (% 로 시작하는 줄은 주석)")
    parser.add_argument("-o", "--out", required=True, help="확장자 없는 출력 경로 (예: figures/eq/gain)")
    parser.add_argument("--scale", type=float, default=SLIDE_SCALE,
                        help=f"PNG 크기 — IEEE 본문 10 pt 의 몇 배 (기본 {SLIDE_SCALE} = 16 pt 상당)")
    args = parser.parse_args()
    if not args.latex and not args.file:
        parser.error("식이나 --file 을 준다")
    latex = args.file.read_text(encoding="utf-8") if args.file else args.latex
    try:
        png = render(latex, args.out, args.scale)
    except (ValueError, RuntimeError) as e:
        print(e)
        sys.exit(1)
    svg = png.with_suffix(".svg").read_text(encoding="utf-8")
    w, h = (float(v) for v in re.findall(r'(?:width|height)="([\d.]+)pt"', svg)[:2])
    print(f"{png.with_suffix('')} — SVG {w:.0f} × {h:.0f} pt (본문 10 pt), PNG {args.scale:g} 배, .tex")


if __name__ == "__main__":
    main()
