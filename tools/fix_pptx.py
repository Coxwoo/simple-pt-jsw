"""다른 도구 (AI 슬라이드 내보내기·온라인 편집기 등) 로 만든 pptx 를 이 양식에 맞게 고친다.

1. 한글 글꼴: 글자마다 적힌 한글 (ea) 글꼴을 "Noto Sans KR" 로 바꾼다 (테마의 빈 ea 도).
   영문 글꼴만 적힌 pptx 는 PowerPoint 가 한글을 다른 글꼴 (맑은 고딕 등) 로 바꿔 그린다.
2. 한글 줄바꿈: 한글이 든 run 의 언어를 ko-KR 로 바꾼다. en-US 로 적힌 한글은 PowerPoint 가 낱말 중간에서 자른다.
3. 아래 첨자 (--sub "V_IN,g_m"): 밑줄로 적힌 기호의 밑줄 뒤를 PowerPoint 아래 첨자로 바꾼다 (장의 글만, 노트는 그대로).
   목록에 없는 밑줄 낱말 (net·파일 이름 등) 은 그대로 두고 끝에 알린다.
4. (--strip-notes) 제출본: 발표자 노트의 글을 지운다 — 남에게 내는 파일에 메모를 남기지 않는다.

내용·배치는 바꾸지 않는다. 원본은 그대로 두고 옆에 `<이름>_fixed.pptx` (노트 있음) 나
`<이름>_submit.pptx` (--strip-notes, 노트 없음) 를 쓴다.
simple_pt.py 로 만든 pptx 는 1·2 가 이미 되어 있어 --strip-notes 만 쓰면 된다.

    python tools/fix_pptx.py deck.pptx
    python tools/fix_pptx.py deck.pptx --sub "V_IN,V_OUT,g_m" --strip-notes
"""

import argparse
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

KR = "Noto Sans KR"
UNDERSCORE_RE = re.compile(r"(?<!\w)[^\W\d_]\w*_\w+")
RUN_RE = re.compile(r"<a:r>(<a:rPr\b[^>]*?/>|<a:rPr\b[^>]*?>.*?</a:rPr>)?(<a:t(?:\s[^>]*)?>)(.*?)</a:t></a:r>", re.S)
HANGUL_RE = re.compile(r"[가-힣]")


def fix_ea(xml):
    return re.subn(r'<a:ea typeface="[^"]*"', f'<a:ea typeface="{KR}"', xml)


def with_baseline(rpr):
    """run 속성에 아래 첨자 (baseline −25000) 를 넣는다. 속성이 없으면 새로 만든다."""
    if not rpr:
        return '<a:rPr baseline="-25000"/>'
    m = re.match(r"<a:rPr\b([^>]*?)(/?>)", rpr)       # 속성과 닫는 꼴 (/> 또는 >) 을 나눈다
    attrs, close = m.group(1), m.group(2)
    if 'baseline="' in attrs:
        attrs = re.sub(r'baseline="[^"]*"', 'baseline="-25000"', attrs)
    else:
        attrs += ' baseline="-25000"'
    return f"<a:rPr{attrs}{close}" + rpr[m.end():]


def add_subscripts(xml, token_re):
    """run 마다 기호를 찾아 밑줄 앞은 같은 run 에, 뒤는 아래 첨자 run 으로 나눈다.
    돌려주는 것: 고친 xml, 바꾼 기호 Counter, 그대로 둔 밑줄 낱말 Counter."""
    done, left = Counter(), Counter()

    def one_run(m):
        rpr, t_open, text = m.group(1) or "", m.group(2), m.group(3)
        rest = token_re.sub("", text) if token_re else text
        left.update(UNDERSCORE_RE.findall(rest))
        if not token_re or not token_re.search(text):
            return m.group(0)
        runs, plain, pos = [], "", 0
        for tm in token_re.finditer(text):
            base, sub = tm.group(1).split("_", 1)
            plain += text[pos:tm.start()] + base
            runs.append(f"<a:r>{rpr}{t_open}{plain}</a:t></a:r>")
            runs.append(f"<a:r>{with_baseline(rpr)}{t_open}{sub}</a:t></a:r>")
            done[tm.group(1)] += 1
            plain, pos = "", tm.end()
        if text[pos:]:
            runs.append(f"<a:r>{rpr}{t_open}{text[pos:]}</a:t></a:r>")
        return "".join(runs)

    return RUN_RE.sub(one_run, xml), done, left


def tag_korean(xml):
    """한글이 든 run 의 언어를 ko-KR 로 바꾼다 (속성이 없으면 넣는다). 돌려주는 것: 고친 xml, 바꾼 run 수."""
    count = 0

    def one_run(m):
        nonlocal count
        rpr, t_open, text = m.group(1) or "", m.group(2), m.group(3)
        if not HANGUL_RE.search(text) or 'lang="ko-KR"' in rpr:
            return m.group(0)
        count += 1
        if not rpr:
            rpr = '<a:rPr lang="ko-KR"/>'
        elif 'lang="' in rpr.split(">", 1)[0]:
            rpr = re.sub(r'\blang="[^"]*"', 'lang="ko-KR"', rpr, count=1)
        else:
            rpr = rpr[:6] + ' lang="ko-KR"' + rpr[6:]
        return f"<a:r>{rpr}{t_open}{text}</a:t></a:r>"

    return RUN_RE.sub(one_run, xml), count


def strip_notes(xml):
    """노트 본문 (ph type="body") 의 글만 지운다 — 노트 쪽과 그 안의 장 그림은 남긴다."""
    def one_sp(m):
        sp = m.group(0)
        if 'type="body"' not in sp:
            return sp
        return re.sub(r"<p:txBody>.*</p:txBody>", "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>", sp, flags=re.S)
    return re.sub(r"<p:sp>.*?</p:sp>", one_sp, xml, flags=re.S)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="pptx 의 한글 글꼴·줄바꿈·아래 첨자를 고치고, 노트를 지운 제출본을 만든다.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--sub", default="", help='아래 첨자로 바꿀 기호, 쉼표로 (예: "V_IN,g_m") — 밑줄 뒤가 아래 첨자')
    parser.add_argument("--strip-notes", action="store_true", help="발표자 노트의 글을 지운 제출본을 만든다")
    args = parser.parse_args()
    tokens = [t.strip() for t in args.sub.split(",") if t.strip()]
    token_re = (re.compile(r"(?<!\w)(" + "|".join(map(re.escape, sorted(tokens, key=len, reverse=True))) + r")(?!\w)")
                if tokens else None)
    src = args.pptx
    dst = src.with_name(f"{src.stem}_{'submit' if args.strip_notes else 'fixed'}.pptx")
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    out = {n: zin.read(n) for n in names}
    ea_total = notes_total = ko_slides = ko_notes = 0
    sub_done, sub_left = Counter(), Counter()
    for n in names:
        if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n):
            xml = out[n].decode("utf-8")
            if args.strip_notes:
                xml = strip_notes(xml)
                notes_total += 1
            xml, k = tag_korean(xml)
            ko_notes += k
            out[n] = xml.encode("utf-8")
        if re.match(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster|theme/theme)\d+\.xml$", n):
            xml, k = fix_ea(out[n].decode("utf-8"))
            ea_total += k
            if re.match(r"ppt/slides/slide\d+\.xml$", n):
                xml, done, left = add_subscripts(xml, token_re)
                sub_done.update(done)
                sub_left.update(left)
                xml, k2 = tag_korean(xml)
                ko_slides += k2
            out[n] = xml.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(zin.getinfo(n), out[n])
    notes_txt = f", 노트 {notes_total} 장 비움" if args.strip_notes else ""
    print(f"{dst.name}: 한글 글꼴 칸 {ea_total} 곳 → {KR}, 아래 첨자 {sum(sub_done.values())} 곳{notes_txt}")
    print(f"  한글 run → ko-KR: 장 {ko_slides} 곳, 노트 {ko_notes} 곳")
    if sub_done:
        print("  아래 첨자: " + ", ".join(f"{k} {v}" for k, v in sub_done.most_common()))
    if sub_left:
        print("  그대로 둔 밑줄 낱말 (--sub 에 없음): " + ", ".join(f"{k} {v}" for k, v in sub_left.most_common()))


if __name__ == "__main__":
    main()
