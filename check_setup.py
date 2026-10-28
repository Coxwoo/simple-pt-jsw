"""처음 쓰는 PC 에서 simple-pt 에 필요한 것을 확인한다 — Python 패키지, 글꼴, figstyle, 장을 그림으로 보는 도구.

    python check_setup.py              # 확인만 한다
    python check_setup.py --install    # 빠진 Python 패키지를 설치하고 (requirements.txt) matplotlib 글꼴 목록을 새로 만든다

글꼴은 자동으로 설치하지 않는다 — 아래 안내 (공식 배포처) 대로 설치한 뒤 --install 로 글꼴 목록을 새로 만들고 다시 확인한다.
끝 코드: 꼭 필요한 것이 모두 있으면 0, 아니면 1.
"""
import argparse
import importlib
import platform
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PACKAGES = [  # (import 이름, pip 이름, 최소 판, 쓰는 곳)
    ("pptx", "python-pptx", "1.0", "pptx 만들기 (tools/simple_pt.py)"),
    ("PIL", "Pillow", "9.0", "그림 크기 읽기"),
    ("lxml", "lxml", "4.9", "pptx 속 XML"),
    ("matplotlib", "matplotlib", "3.7", "식 그림 (tools/equations.py), 그래프 (figstyle)"),
    ("numpy", "numpy", "1.24", "그래프 (figstyle)"),
]
OPTIONAL = [("sympy", "sympy", "1.12", "식 검산 (선택)")]
FONTS = [  # (글꼴, 꼭 필요한가, 쓰는 곳, 없을 때)
    ("Arial", True, "슬라이드·그래프의 영문", "비슷한 글꼴로 바뀌어 글 폭이 달라진다"),
    ("Arial Narrow", False, "그래프 안 글자 (figstyle)", "Arial 로 대신 — 그래프 안 글자가 조금 넓어진다"),
    ("Noto Sans KR", True, "슬라이드의 한글", "맑은 고딕 등으로 바뀌어 모양·줄바꿈이 달라진다"),
]
FONT_HELP = {
    "Arial": "Windows·macOS 에 기본으로 있다. Linux 는 같은 폭의 대체 글꼴: sudo apt install fonts-liberation",
    "Arial Narrow": "Microsoft Office 와 함께 설치된다. Linux 는 fonts-liberation 의 Liberation Sans Narrow",
    "Noto Sans KR": ("https://fonts.google.com/noto/specimen/Noto+Sans+KR → Get font → Download all → "
                     "압축 속 static/ 의 NotoSansKR-Regular.ttf·NotoSansKR-Bold.ttf 를 설치 "
                     "(Windows: 오른쪽 클릭 → 설치, macOS: 서체 관리자, Linux: ~/.local/share/fonts 에 복사 후 fc-cache -f)"),
}
FIGSTYLE_PLACES = [HERE.parent / "figstyle-jsw", HERE.parent / "figstyle", Path.home() / ".claude" / "skills" / "figstyle"]
FIGSTYLE_URL = "https://github.com/Coxwoo/figstyle-jsw.git"
OK, NO, SOSO = "✓", "✗", "△"


def _version_ok(have, need):
    def nums(v):
        out = []
        for part in str(v).split("."):
            digits = "".join(ch for ch in part if ch.isdigit())
            if not digits:
                break
            out.append(int(digits))
        return out
    return nums(have) >= nums(need)


def check_packages(items):
    rows, missing = [], []
    for mod, pip_name, need, use in items:
        try:
            ver = getattr(importlib.import_module(mod), "__version__", "?")
            good = ver == "?" or _version_ok(ver, need)
            rows.append((OK if good else NO, f"{pip_name} {ver}", use + ("" if good else f" — {need} 이상 필요")))
            if not good:
                missing.append(pip_name)
        except ImportError:
            rows.append((NO, pip_name, f"{use} — 없음"))
            missing.append(pip_name)
    return rows, missing


def check_fonts():
    """설치된 글꼴 파일을 직접 읽어 본다 (matplotlib 글꼴 목록과도 비교)."""
    try:
        from matplotlib import font_manager as fm
    except ImportError:
        return [(NO, "글꼴", "matplotlib 을 설치한 뒤 다시 확인")], [], False
    files = {}
    for entry in fm.FontManager().ttflist:                     # 지금 설치된 글꼴을 새로 읽는다 (캐시에는 쓰지 않는다)
        files.setdefault(entry.name, []).append((Path(entry.fname).name, entry.weight))
    cached = {f.name for f in fm.fontManager.ttflist}
    rows, missing_required, stale = [], [], False
    for name, required, use, fallback in FONTS:
        found = files.get(name, [])
        if not found:
            rows.append((NO if required else SOSO, name, f"{use} — 없음 ({fallback})"))
            if required:
                missing_required.append(name)
            continue
        variable_only = all("VF" in fn or "[" in fn for fn, _ in found)
        if variable_only:
            rows.append((SOSO, name, f"{use} — 가변 판 (VF) 만 있음: PowerPoint 는 되지만 그래프는 가늘게 나옴 → static 판 설치 권장"))
        else:
            rows.append((OK, name, use))
        if name not in cached:
            stale = True
    return rows, missing_required, stale


def check_figstyle():
    for place in FIGSTYLE_PLACES:
        if (place / "figstyle.py").is_file():
            return (OK, "figstyle", f"그래프 스타일 — {place}")
    return (SOSO, "figstyle", f"그래프 스타일 — 못 찾음: git clone {FIGSTYLE_URL} ~/.claude/skills/figstyle")


def check_renderer():
    """장을 그림으로 뽑아 눈으로 볼 도구 — PowerPoint (Windows) 나 LibreOffice. 없으면 PowerPoint 에서 직접 연다."""
    if platform.system() == "Windows":
        try:
            import winreg
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application"))
            return (OK, "PowerPoint", "장을 그림으로 뽑아 확인 (COM)")
        except OSError:
            pass
    for cand in ("soffice", "libreoffice", r"C:\Program Files\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if shutil.which(cand) or Path(cand).is_file():
            return (OK, "LibreOffice", "장을 그림으로 뽑아 확인 (soffice --headless --convert-to pdf)")
    return (SOSO, "PowerPoint·LibreOffice", "없음 — 장을 그림으로 확인하려면 둘 중 하나 (선택)")


def refresh_font_cache():
    try:
        from matplotlib import font_manager as fm
        fm._load_fontmanager(try_read_cache=False)
        print("  matplotlib 글꼴 목록을 새로 만들었다")
    except Exception as e:
        print(f"  matplotlib 글꼴 목록을 새로 만들지 못함 ({e}) — matplotlib 캐시 폴더의 fontlist-*.json 을 지운다")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="simple-pt 에 필요한 것을 확인한다")
    parser.add_argument("--install", action="store_true", help="빠진 Python 패키지를 설치하고 글꼴 목록을 새로 만든다")
    args = parser.parse_args()

    print(f"Python {platform.python_version()} ({platform.system()})")
    if sys.version_info < (3, 9):
        print(f"  {NO} Python 3.9 이상이 필요하다")
    rows, missing = check_packages(PACKAGES)
    opt_rows, _ = check_packages(OPTIONAL)
    if args.install and missing:
        print(f"\n설치: {', '.join(missing)}")
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(HERE / "requirements.txt")], check=False)
        importlib.invalidate_caches()
        rows, missing = check_packages(PACKAGES)
    font_rows, missing_fonts, stale = check_fonts()
    if args.install and stale:
        refresh_font_cache()
        font_rows, missing_fonts, stale = check_fonts()

    print("\n[Python 패키지]")
    for mark, what, note in rows + opt_rows:
        print(f"  {mark} {what:<22} {note}")
    print("\n[글꼴]")
    for mark, what, note in font_rows:
        print(f"  {mark} {what:<22} {note}")
    print("\n[함께 쓰는 것]")
    for mark, what, note in (check_figstyle(), check_renderer()):
        print(f"  {mark} {what:<22} {note}")

    todo = []
    if missing:
        todo.append("python check_setup.py --install   # Python 패키지 설치")
    for name in missing_fonts + [r[1] for r in font_rows if r[0] == SOSO and r[1] in FONT_HELP]:
        todo.append(f"{name}: {FONT_HELP[name]}")
    if stale:
        todo.append("글꼴을 설치했으면: python check_setup.py --install   # matplotlib 글꼴 목록 새로 만들기")
    print("\n" + ("할 일:\n  - " + "\n  - ".join(dict.fromkeys(todo)) if todo else "모두 준비됨"))
    sys.exit(1 if (missing or missing_fonts or sys.version_info < (3, 9)) else 0)


if __name__ == "__main__":
    main()
