"""simple_pt — "Simple is best" 양식의 발표 자료 (pptx) 를 만드는 작은 모듈 (python-pptx).

디자인 값 (자리·글꼴·크기·색) 이 이 파일 한 곳에 있어서, 누가 어떤 AI 로 만들어도 같은 모양이 나온다.
규칙은 rules/ — design (모양), writing (글), structure (덱 구성), checklist (검수).

쓰는 법:
    from simple_pt import Deck
    d = Deck("Demo", parts=["Why", "Design", "Result"])     # parts 를 주면 오른쪽 위에 진행 표시 (선택)
    d.cover("Short Topic:\\nWhat It Does", subtitle="한 줄 부제", kicker="Team Seminar",
            name="Your Name", affiliation="Dept. · Lab", date="YYYY-MM-DD")
    d.outline([("Why", "한 줄 설명"), ("Design", "…"), ("Result", "…")])
    d.divider("Result")                                       # 파트 구분 장 — 가운데 제목 + 밑줄
    d.panels("Results", "이 장의 질문이나 결론 한 문장", part="Result",
             panels=[{"label": "Amp alone: gain (sim)", "image": "figures/gain.png",
                      "takeaway": "굵은 결론 한 줄", "caption": "회색 조건 한 줄"}],
             key="빨간 결론 한 줄 — 파트를 맺는 장에만", notes="말할 대본")
    d.save("deck.pptx")                # 규칙에 어긋나 보이는 곳은 경고로 알린다

장 종류 (레이아웃 이름은 rules/design.md 7 절):
    cover           표지
    outline         목차 — 파트 이름 + 한 줄 설명 (번호 없음)
    divider         파트 구분 장 — 가운데 제목 + 밑줄 (그림·쪽 번호·진행 표시 없음)
    panels          그림 1–3 칸 (칸마다 이름표·그림·결론 줄·캡션)                              F
    figure_bullets  왼쪽 그림 + 오른쪽 불릿 3–5 개                                             A
    equation        식 (채움 없는 검정 2¼ pt 상자 — box=False 면 상자 없이) + 아래 그림·글 묶음   B
    cards           글 묶음 2–3 개 (소제목 + 불릿, 상자 없음) — 그림 없는 장이 되니 경고          E·G
    table           논문식 표 (채움·세로선 없음, highlight 칸만 빨강)                            C
    summary         요약 — 필요할 때만                                                          H
    appendix_index  Appendix 첫 장 — 질문 | Appendix | 장 (장 번호는 save 때 채운다)
appendix_index 뒤의 장은 쪽 번호가 A1, A2 … 가 된다.
내용 장은 key= (빨간 줄 — 파트를 맺는 장에만), part= (진행 표시), refs= (아래 작은 글씨 참고문헌) 를 받는다.

글 표시: **굵게**, V_{IN} 아래 첨자, x^{2} 위 첨자. 줄바꿈은 "\\n", 불릿은 리스트. τ 는 저절로 Times New Roman 기울임.
"""
import datetime
import math
import re
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# ── 디자인 값 — 바꾸면 rules/design.md 도 같이 고친다 ─────────────────────────
W, H = 1920, 1080                  # 캔버스 px (= 13.333 × 7.5 in, 1 pt = 2 px)
MX, TOP, BOTTOM = 128, 72, 920     # 좌우 여백, 제목 위 끝, 본문 아래 끝
CW = W - 2 * MX                    # 본문 폭 1664
PAGE_Y = H - 64                    # 쪽 번호·표지 날짜·참고문헌 줄의 아래 끝
PROG_W = 536                       # 오른쪽 위 진행 표시 칸 폭 — 진행 표시가 있으면 제목 칸이 그만큼 (+32) 줄어든다

LATIN, KR, GREEK = "Arial", "Noto Sans KR", "Times New Roman"   # GREEK: τ (Arial 의 τ 는 작은 t 처럼 보인다)
INK, TEXT, GREY, MUTED, RED = "111111", "1F1F1F", "4A4A4A", "767676", "B4161E"
BOX_LINE, RULE = "000000", "D9D9D9"                 # 상자 (채움 없음) 의 선, 얇은 구분선
BOX_W, TABLE_EDGE, TABLE_HEAD = 2.25, 1.0, 0.75     # 선 굵기 pt — 상자, 논문식 표의 위·맨 아래 선, 머리 줄 아래 선

SIZE = {  # 글자 크기 px (pt 의 2 배)
    "title": 52, "subtitle": 40, "label": 28, "body": 32, "card_title": 32, "card_body": 28,
    "table": 28, "takeaway": 32, "caption": 24, "key": 34, "page": 24, "notes": 24, "progress": 24, "ref": 16,
    "cover_title": 72, "cover_kicker": 24, "cover_name": 32, "divider": 80, "outline_part": 44, "outline_desc": 36,
}
LINE = {  # 줄 높이 — 글자 크기의 배수 (CSS line-height 와 같은 뜻, pptx 에는 정확한 pt 로 넣는다)
    "title": 1.15, "subtitle": 1.3, "label": 1.3, "body": 1.45, "card": 1.4, "table": 1.3, "takeaway": 1.3,
    "caption": 1.4, "key": 1.3, "page": 1.2, "progress": 1.2, "ref": 1.3, "cover_title": 1.12, "cover": 1.3,
    "divider": 1.15, "outline": 1.3,
}
NO_TABLE_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # PowerPoint "스타일 없음, 눈금 없음"

HANGUL = re.compile(r"[ᄀ-ᇿ㄰-㆏가-힣]")
MARK = re.compile(r"\*\*(.+?)\*\*|_\{(.+?)\}|\^\{(.+?)\}")
SENTENCE_END = re.compile(r"(다|요)[.!]?$")                   # "~다." 로 끝나는 문장형
TAU = re.compile(r"(τ)")


def px(v):
    return Emu(round(v * 6350))                                # 1 px = 6350 EMU


# ── 글 ──────────────────────────────────────────────────────────────────────
def _plain(text):
    return MARK.sub(lambda m: next(g for g in m.groups() if g is not None), str(text))


def _runs(text):
    """글 표시를 run 조각으로 — [(글, 굵게, baseline)]."""
    out, pos = [], 0
    for m in MARK.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False, 0))
        bold, sub, sup = m.groups()
        if bold is not None:
            out.append((bold, True, 0))
        elif sub is not None:
            out.append((sub, False, -25000))
        else:
            out.append((sup, False, 30000))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False, 0))
    return out or [("", False, 0)]


def _width(text, size, bold=False):
    """한 줄 글의 폭 어림 (px) — 한글·한자 1.0 em, 영문 대문자 0.68, 그 밖 0.55, 빈칸 0.28."""
    w = 0.0
    for ch in _plain(text):
        if HANGUL.match(ch) or ord(ch) >= 0x2E80:
            w += 1.0
        elif ch == " ":
            w += 0.28
        elif ch.isupper() or ch in "@%":
            w += 0.68
        else:
            w += 0.55
    return w * size * (1.07 if bold else 1.0)


def _text_px(text, size, bold=False):
    """한 줄 글의 폭 (px) — Arial 을 찾으면 실제 글꼴로 재고, 없으면 _width 어림."""
    from PIL import ImageFont
    for name in (("arialbd.ttf", "Arial Bold.ttf") if bold else ("arial.ttf", "Arial.ttf")):
        try:
            return ImageFont.truetype(name, size).getlength(_plain(text))
        except OSError:
            continue
    return _width(text, size, bold)


def _lines(text, size, width, bold=False):
    """줄바꿈까지 어림한 줄 수."""
    if not text:
        return 0
    paras = text if isinstance(text, (list, tuple)) else str(text).split("\n")
    return sum(max(1, math.ceil(_width(p, size, bold) / width)) for p in paras)


def _font(run, size, color, bold=False, baseline=0, caps=False, spacing=0, greek=False):
    """run 의 글꼴 — 영문 Arial, 한글 (ea) Noto Sans KR, 언어는 한글이 있으면 ko-KR (줄바꿈이 낱말 단위가 된다).
    greek: τ 한 글자 run — Times New Roman 기울임 (Arial 의 τ 는 작은 t 처럼 보인다)."""
    f = run.font
    f.size = Pt(size / 2)
    f.bold = bold
    f.color.rgb = RGBColor.from_string(color)
    f.name = GREEK if greek else LATIN
    if greek:
        f.italic = True
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        old = rpr.find(qn(tag))
        if old is not None:
            rpr.remove(old)
    ea, cs = etree.Element(qn("a:ea")), etree.Element(qn("a:cs"))
    ea.set("typeface", KR)
    cs.set("typeface", GREEK if greek else LATIN)
    latin = rpr.find(qn("a:latin"))
    latin.addnext(ea)
    ea.addnext(cs)
    rpr.set("lang", "ko-KR" if HANGUL.search(run.text) else "en-US")
    if baseline:
        rpr.set("baseline", str(baseline))
    if caps:
        rpr.set("cap", "all")
    if spacing:
        rpr.set("spc", str(spacing))


def _bullet(p, size):
    ppr = p._p.get_or_add_pPr()
    ind = int(px(size * 0.85))
    ppr.set("marL", str(ind))
    ppr.set("indent", str(-ind))
    bu_font = etree.SubElement(ppr, qn("a:buFont"))
    bu_font.set("typeface", LATIN)
    bu_char = etree.SubElement(ppr, qn("a:buChar"))
    bu_char.set("char", "•")


def _fill_frame(tf, text, size, color, bold=False, line=1.3, align=PP_ALIGN.LEFT, bullets=False,
                gap_after=0, caps=False, spacing=0):
    paras = list(text) if isinstance(text, (list, tuple)) else str(text).split("\n")
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = Pt(size / 2 * line)
        if gap_after and i < len(paras) - 1:
            p.space_after = Pt(gap_after / 2)
        if bullets:
            _bullet(p, size)
        for part, b, base in _runs(str(para)):
            for piece in [x for x in TAU.split(part) if x] or [""]:
                r = p.add_run()
                r.text = piece
                _font(r, size, color, bold or b, base, caps, spacing, greek=(piece == "τ"))


def _textbox(slide, x, y, w, h, text, size, color, bold=False, line=1.3, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, bullets=False, gap_after=0, caps=False, spacing=0):
    box = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    _fill_frame(tf, text, size, color, bold, line, align, bullets, gap_after, caps, spacing)
    return box


# ── 도형·그림·표 ─────────────────────────────────────────────────────────────
def _plain_shape(shape):
    """테마 스타일 (그림자·테마 색) 을 떼어 낸다 — 채움·선은 직접 준다."""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _rect(slide, x, y, w, h, fill=None, line=None, radius=0, width=0.75):
    """사각형 — fill 이 없으면 채움 없음, line 이 없으면 선 없음. width 는 선 굵기 (pt)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                 px(x), px(y), px(w), px(h))
    _plain_shape(shp)
    if radius:
        shp.adjustments[0] = radius / min(w, h)
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = RGBColor.from_string(line)
        shp.line.width = Pt(width)
    else:
        shp.line.fill.background()
    return shp


def _rule(slide, x1, y, x2, color=RULE):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, px(x1), px(y), px(x2), px(y))
    _plain_shape(ln)
    ln.line.color.rgb = RGBColor.from_string(color)
    ln.line.width = Pt(0.75)


def _fixed_zip_dates(path):
    """pptx (zip) 안 파일의 날짜를 고정한다 — 만든 날이 파일에 남지 않고, 같은 입력이면 같은 파일이 된다."""
    with zipfile.ZipFile(path) as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            zi = zipfile.ZipInfo(info.filename, date_time=(2000, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)


def _image_size(path):
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"그림이 없다: {path}")
    with Image.open(path) as im:
        return im.size


def _equation_png(eq):
    """식 그림 — PNG 경로면 그대로, LaTeX 글이면 tools/equations.py 로 임시 폴더에 만든다."""
    if isinstance(eq, (str, Path)) and str(eq).lower().endswith(".png"):
        if not Path(eq).is_file():
            raise FileNotFoundError(f"식 그림이 없다: {eq}")
        return Path(eq)
    import hashlib
    import sys
    import tempfile
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from equations import render
    name = hashlib.sha1(repr(eq).encode("utf-8")).hexdigest()[:12]
    return render(eq, Path(tempfile.gettempdir()) / "simple_pt_eq" / name)


def _picture(slide, path, x, y, w, h, alt=""):
    pic = slide.shapes.add_picture(str(path), px(x), px(y), px(w), px(h))
    pic._element.nvPicPr.cNvPr.set("descr", alt or Path(path).stem.replace("_", " "))
    return pic


def _cell(cell, text, size, bold, color, borders):
    """표 칸 — 채움 없음. borders: (위, 아래) 선 굵기 pt (0 이면 선 없음). 왼쪽·오른쪽 선은 없다."""
    cell.margin_left = cell.margin_right = px(16)
    cell.margin_top = cell.margin_bottom = px(10)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.fill.background()
    tf = cell.text_frame
    tf.word_wrap = True
    _fill_frame(tf, text, size, color, bold, LINE["table"])
    tcpr = cell._tc.get_or_add_tcPr()
    top, bottom = borders
    for i, (tag, w) in enumerate((("a:lnL", 0), ("a:lnR", 0), ("a:lnT", top), ("a:lnB", bottom))):
        old = tcpr.find(qn(tag))                                   # 칸 테두리 — 채움보다 앞에 둔다
        if old is not None:
            tcpr.remove(old)
        if w:
            ln = etree.Element(qn(tag), w=str(round(w * 12700)), cap="flat", cmpd="sng", algn="ctr")
            clr = etree.SubElement(etree.SubElement(ln, qn("a:solidFill")), qn("a:srgbClr"))
            clr.set("val", INK)
            etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
        else:
            ln = etree.Element(qn(tag), w="0")
            etree.SubElement(ln, qn("a:noFill"))
        tcpr.insert(i, ln)


def _table(slide, x, y, w, header, rows, widths=None, size=SIZE["table"], highlight=()):
    """논문식 표 — 칸 채움·세로선 없이 위 1 pt·머리 줄 아래 0.75 pt·맨 아래 1 pt 검은 선, 머리 줄은 굵게.
    highlight: 빨강으로 짚을 칸 [(줄, 칸)] — 줄은 rows 의 번호 (0 부터). 돌려주는 것: python-pptx 표, 높이 (px)."""
    ncol = len(header)
    widths = widths or [1] * ncol
    col_w = [w * v / sum(widths) for v in widths]
    line_h = size * LINE["table"]
    grid = [list(map(str, header))] + [list(map(str, r)) for r in rows]
    heights = [20 + line_h * max(1, max(_lines(c, size, cw - 32, bold=(r == 0)) for c, cw in zip(cells, col_w)))
               for r, cells in enumerate(grid)]
    frame = slide.shapes.add_table(len(grid), ncol, px(x), px(y), px(w), px(sum(heights)))
    tbl = frame.table
    tblpr = tbl._tbl.tblPr
    tblpr.set("firstRow", "0")
    tblpr.set("bandRow", "0")
    sid = tblpr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblpr, qn("a:tableStyleId"))
    sid.text = NO_TABLE_STYLE
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = px(cw)
    for r, hgt in enumerate(heights):
        tbl.rows[r].height = px(hgt)
    last, hot = len(grid) - 1, {(r + 1, c) for r, c in highlight}
    for r, cells in enumerate(grid):
        top = TABLE_EDGE if r == 0 else (TABLE_HEAD if r == 1 else 0)
        bottom = TABLE_HEAD if r == 0 else (TABLE_EDGE if r == last else 0)
        for c, val in enumerate(cells):
            red = (r, c) in hot
            _cell(tbl.cell(r, c), val, size, bold=(r == 0 or red), color=RED if red else (INK if r == 0 else TEXT),
                  borders=(top, bottom))
    return tbl, sum(heights)


# ── 덱 ──────────────────────────────────────────────────────────────────────
class Deck:
    """장을 순서대로 더하고 save() 로 쓴다. 장마다 notes 에 말할 대본을 넣는다.
    parts: 파트 이름 (예: ["Why", "Design", "Result"]) — 주면 장마다 part= 로 오른쪽 위 진행 표시를 단다."""

    def __init__(self, title="", parts=()):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = px(W), px(H)
        self.title = title
        self.parts = list(parts)
        self.warnings = []
        self._count = 0                 # 장 수
        self._appendix = None           # appendix_index 뒤로 몇 번째 A 장인가
        self._labels = {}               # 쪽 번호 글 → 장 번호
        self._page = {}                 # slide_id → 쪽 번호 글
        self._index_cells = []          # (표 칸, 가리키는 쪽 번호 글)
        self._kinds = []                # (장 번호, 꼴) — 같은 꼴이 이어지는지 save() 가 본다
        self._body = 0                  # Appendix 앞 내용 장 수
        self._keys = 0                  # 그 가운데 빨간 줄이 있는 장 수

    # 공통 ------------------------------------------------------------------
    def _warn(self, page, msg):
        self.warnings.append(f"[{page}] {msg}")

    def _new_slide(self, kind):
        self._count += 1
        self._kinds.append((self._count, kind))
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])     # 빈 레이아웃

    def _start(self, title, subtitle, key, kind, page=None, part=None):
        s = self._new_slide(kind)
        if page is None:
            if self._appendix is None:
                page = str(self._count)
            else:
                self._appendix += 1
                page = f"A{self._appendix}"
        self._labels[page] = self._count
        self._page[s.slide_id] = page
        if self._appendix is None and kind != "outline":
            self._body += 1
            self._keys += bool(key)
        elif key:
            self._warn(page, "Appendix 장에는 빨간 줄을 두지 않는다 — 결론은 그림 아래 굵은 줄로")
        tw = CW - (PROG_W + 32 if self._progress(s, part, page) else 0)
        self._check_text(page, title, subtitle, key, tw)
        th = SIZE["title"] * LINE["title"]
        _textbox(s, MX, TOP, tw, th, title, SIZE["title"], INK, bold=True, line=LINE["title"])
        top = TOP + th + 12
        if subtitle:
            n = _lines(subtitle, SIZE["subtitle"], CW)
            sh = n * SIZE["subtitle"] * LINE["subtitle"]
            _textbox(s, MX, top, CW, sh, subtitle, SIZE["subtitle"], GREY, line=LINE["subtitle"])
            top += sh
        top += 18 + 40
        bottom = BOTTOM
        if key:
            kh = _lines(key, SIZE["key"], CW, bold=True) * SIZE["key"] * LINE["key"]
            _textbox(s, MX, BOTTOM - kh, CW, kh, key, SIZE["key"], RED, bold=True, line=LINE["key"])
            bottom = BOTTOM - kh - 40
        return s, top, bottom

    def _progress(self, s, part, page):
        """오른쪽 위 진행 표시 — 파트 이름을 › 로 잇고 지금 파트만 굵은 검정. Appendix 는 'Appendix › 파트'.
        Deck(parts=…) 가 없거나 본문 장에 part 가 없으면 달지 않는다. 돌려주는 것: 달았는지."""
        if not self.parts:
            return False
        if self._appendix is None:
            if part is None:
                return False
            if part not in self.parts:
                self._warn(page, f"part '{part}' 가 Deck(parts=…) 에 없음")
            items = [(p, p == part) for p in self.parts]
        else:
            items = [("Appendix", not part)] + ([(part, True)] if part else [])
        h = SIZE["progress"] * LINE["progress"]
        box = s.shapes.add_textbox(px(W - MX - PROG_W), px(TOP + 12), px(PROG_W), px(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        p.line_spacing = Pt(SIZE["progress"] / 2 * LINE["progress"])
        for i, (name, now) in enumerate(items):
            for text, color, bold in ([(" › ", MUTED, False)] if i else []) + [(name, INK if now else MUTED, now)]:
                r = p.add_run()
                r.text = text
                _font(r, SIZE["progress"], color, bold)
        return True

    def _finish(self, s, notes, refs=()):
        _textbox(s, W - MX - 240, PAGE_Y - SIZE["page"] * LINE["page"], 240, SIZE["page"] * LINE["page"],
                 self._page[s.slide_id], SIZE["page"], MUTED, line=LINE["page"], align=PP_ALIGN.RIGHT,
                 anchor=MSO_ANCHOR.BOTTOM)
        if refs:                                       # 인용 [n] 의 참고문헌 — 쪽 번호 줄 왼쪽에 작은 회색 글씨
            text = "   ".join(refs)
            rw = CW - 240 - 32
            n = _lines(text, SIZE["ref"], rw)
            if n > 2:
                self._warn(self._page[s.slide_id], f"참고문헌 줄이 {n} 줄 — 두 줄 이하로 (줄여 쓰거나 Appendix 로)")
            rh = n * SIZE["ref"] * LINE["ref"]
            _textbox(s, MX, PAGE_Y - rh, rw, rh, text, SIZE["ref"], MUTED, line=LINE["ref"], anchor=MSO_ANCHOR.BOTTOM)
        self._notes(s, notes)

    def _notes(self, s, notes):
        if notes:
            _fill_frame(s.notes_slide.notes_text_frame, notes, SIZE["notes"], TEXT, line=1.4)

    def _check_text(self, page, title, subtitle, key, tw=CW):
        words = len(_plain(title).split())
        if words > 5:
            self._warn(page, f"제목이 {words} 단어 — 2–4 단어의 짧은 주제로")
        if _lines(title, SIZE["title"], tw, bold=True) > 1:
            self._warn(page, "제목이 한 줄을 넘음")
        for name, t, size, most, bold in (("부제", subtitle, SIZE["subtitle"], 2, False),
                                          ("빨간 줄", key, SIZE["key"], 1, True)):
            if not t:
                continue
            if SENTENCE_END.search(_plain(t).strip()):
                self._warn(page, f"{name}가 문장형 (~다) — 음슴체·단정형으로 (~사용, ~확보, ~이내)")
            if _lines(t, size, CW, bold) > most:
                self._warn(page, f"{name}가 {most} 줄을 넘음 — 줄이거나 노트로")

    def _figure(self, s, x, y, w, h, panel, page):
        """이름표 (있으면) + 그림 + 결론 줄 (굵게) + 캡션 (회색) 한 칸. 그림은 칸 왼쪽에 붙이고 비율을 지킨다.
        돌려주는 것: 그림 폭."""
        label, take, cap = panel.get("label", ""), panel.get("takeaway", ""), panel.get("caption", "")
        iw, ih = _image_size(panel["image"])
        lh = SIZE["label"] * LINE["label"] if label else 0
        img_top = y + (lh + 12 if label else 0)
        take_n, cap_n = _lines(take, SIZE["takeaway"], w, bold=True), _lines(cap, SIZE["caption"], w)
        for _ in range(3):                                # 글 줄 수가 그림 폭에 따라 바뀌어 몇 번 맞춘다
            tkh = take_n * SIZE["takeaway"] * LINE["takeaway"]
            ch = cap_n * SIZE["caption"] * LINE["caption"]
            room = y + h - img_top - (tkh + 12 if take else 0) - (ch + (8 if take else 12) if cap else 0)
            sc = min(w / iw, room / ih)
            pw, ph = iw * sc, ih * sc
            n1, n2 = _lines(take, SIZE["takeaway"], pw, bold=True), _lines(cap, SIZE["caption"], pw)
            if (n1, n2) == (take_n, cap_n):
                break
            take_n, cap_n = n1, n2
        if label:
            _textbox(s, x, y, w, lh, label, SIZE["label"], INK, bold=True, line=LINE["label"])
        _picture(s, panel["image"], x, img_top, pw, ph, panel.get("alt", ""))
        yy = img_top + ph + 12
        if take:
            if take_n > 2:
                self._warn(page, f"그림 아래 결론 줄이 {take_n} 줄 — 한두 줄로")
            if SENTENCE_END.search(_plain(take).strip()):
                self._warn(page, "그림 아래 결론 줄이 문장형 (~다) — 음슴체·단정형으로")
            _textbox(s, x, yy, pw, tkh, take, SIZE["takeaway"], INK, bold=True, line=LINE["takeaway"])
            yy += tkh + 8
        if cap:
            if cap_n > 2:
                self._warn(page, f"캡션이 {cap_n} 줄 — 한두 줄로")
            if SENTENCE_END.search(_plain(cap).strip()):
                self._warn(page, "캡션이 문장형 (~다) — 음슴체·단정형으로")
            _textbox(s, x, yy, pw, ch, cap, SIZE["caption"], GREY, line=LINE["caption"])
        return pw

    # 장 종류 ---------------------------------------------------------------
    def cover(self, title, subtitle="", kicker="", name="", affiliation="", date="", notes=""):
        """표지 — 작은 대문자 회색 줄 → 검정 가로선 → 큰 제목 → 회색 부제 → 이름·소속 → 맨 아래 날짜."""
        s = self._new_slide("cover")
        items = []                                         # (높이, 그리는 함수)
        if kicker:
            kh = SIZE["cover_kicker"] * LINE["cover"]
            items.append((kh, lambda y: _textbox(s, MX, y, CW, kh, kicker, SIZE["cover_kicker"], MUTED, bold=True,
                                                 line=LINE["cover"], caps=True, spacing=200)))
        items.append((6, lambda y: _rect(s, MX, y, 326, 6, INK)))
        tn = _lines(title, SIZE["cover_title"], CW, bold=True)
        th = tn * SIZE["cover_title"] * LINE["cover_title"]
        items.append((th, lambda y: _textbox(s, MX, y, CW, th, title, SIZE["cover_title"], INK, bold=True,
                                             line=LINE["cover_title"])))
        if subtitle:
            sh = _lines(subtitle, SIZE["subtitle"], CW) * SIZE["subtitle"] * LINE["cover"]
            items.append((sh, lambda y: _textbox(s, MX, y, CW, sh, subtitle, SIZE["subtitle"], GREY, line=LINE["cover"])))
        if name or affiliation:
            nh = SIZE["cover_name"] * LINE["cover"]
            hh = nh * 2 + 8 if (name and affiliation) else nh

            def who(y):
                if name:
                    _textbox(s, MX, y, CW, nh, name, SIZE["cover_name"], INK, bold=True, line=LINE["cover"])
                if affiliation:
                    _textbox(s, MX, y + (nh + 8 if name else 0), CW, nh, affiliation, SIZE["cover_name"], MUTED,
                             line=LINE["cover"])
            items.append((hh, who))
        total = sum(h for h, _ in items) + 36 * (len(items) - 1)
        y = 128 + max(0, (H - 128 - 160 - total) / 2)
        for h, draw in items:
            draw(y)
            y += h + 36
        if date:
            dh = SIZE["page"] * LINE["page"]
            _textbox(s, MX, PAGE_Y - dh, CW, dh, date, SIZE["page"], MUTED, line=LINE["page"], anchor=MSO_ANCHOR.BOTTOM)
        if tn > 3:
            self._warn("표지", "제목이 세 줄을 넘음")
        self._notes(s, notes)

    def outline(self, parts=(), title="Outline", subtitle="", notes=""):
        """목차 — 파트 이름 (굵게) + 한 줄 설명, 번호 없이 가는 선으로 나눈다. parts: [(이름, 설명), …]."""
        s, top, bottom = self._start(title, subtitle, "", "outline")
        rows = list(parts)
        rh = min(170, (bottom - top) / max(len(rows), 1))
        nh, dh = SIZE["outline_part"] * LINE["outline"], SIZE["outline_desc"] * LINE["outline"]
        for i, (name, desc) in enumerate(rows):
            y = top + i * rh
            _rule(s, MX, y, MX + CW)
            _textbox(s, MX, y + (rh - nh) / 2, 536, nh, name, SIZE["outline_part"], INK, bold=True,
                     line=LINE["outline"])
            _textbox(s, MX + 560, y + (rh - dh) / 2, CW - 560, dh, desc, SIZE["outline_desc"], GREY,
                     line=LINE["outline"])
        if rows:
            _rule(s, MX, top + len(rows) * rh, MX + CW)
        self._finish(s, notes)

    def divider(self, title, notes=""):
        """파트 구분 장 — 가운데 파트 제목 (굵게) + 그 아래 밑줄. 그림·쪽 번호·진행 표시는 없다."""
        s = self._new_slide("divider")
        th = SIZE["divider"] * LINE["divider"]
        uw = max(160, _text_px(title, SIZE["divider"], bold=True))
        y = (H - (th + 16 + 6)) / 2
        _textbox(s, MX, y, CW, th, title, SIZE["divider"], INK, bold=True, line=LINE["divider"], align=PP_ALIGN.CENTER)
        _rect(s, (W - uw) / 2, y + th + 16, uw, 6, INK)
        if _lines(title, SIZE["divider"], CW, bold=True) > 1:
            self._warn(str(self._count), "구분 장 제목이 한 줄을 넘음")
        self._notes(s, notes)

    def panels(self, title, subtitle="", panels=(), key="", notes="", part=None, refs=()):
        """그림 1–3 칸 — panels: [{"label", "image", "takeaway", "caption", "alt"}].
        label (영어 이름표) 은 무엇을 어떤 조건으로 잰 것인지 보여야 할 때만, takeaway 는 굵은 결론 한 줄, caption 은 회색 조건·세부."""
        n = len(panels)
        s, top, bottom = self._start(title, subtitle, key, f"panels {n}", part=part)
        if not 1 <= n <= 3:
            self._warn(self._page[s.slide_id], f"그림 {n} 칸 — 1–3 칸으로")
        gap = 32 if n >= 3 else 64
        cw = (CW - gap * (n - 1)) / max(n, 1)
        for i, panel in enumerate(panels):
            self._figure(s, MX + i * (cw + gap), top, cw, bottom - top, panel, self._page[s.slide_id])
        self._finish(s, notes, refs)

    def figure_bullets(self, title, subtitle="", image="", label="", caption="", bullets=(), key="", notes="",
                       alt="", fig_width=0.5, takeaway="", part=None, refs=()):
        """왼쪽 그림 (본문 폭의 fig_width, 아래 결론 줄·캡션) + 오른쪽 불릿 3–5 개."""
        s, top, bottom = self._start(title, subtitle, key, "figure_bullets", part=part)
        pw = self._figure(s, MX, top, CW * fig_width, bottom - top,
                          {"image": image, "label": label, "takeaway": takeaway, "caption": caption, "alt": alt},
                          self._page[s.slide_id])
        bx = MX + pw + 64
        bw = MX + CW - bx
        by = top + (SIZE["label"] * LINE["label"] + 12 if label else 0)
        need = (_lines(list(bullets), SIZE["body"], bw - SIZE["body"]) * SIZE["body"] * LINE["body"]
                + 16 * max(len(bullets) - 1, 0))
        if need > bottom - by:
            self._warn(self._page[s.slide_id], "불릿이 칸을 넘음 — 줄이거나 노트로")
        if len(bullets) > 5:
            self._warn(self._page[s.slide_id], f"불릿 {len(bullets)} 개 — 3–5 개로")
        _textbox(s, bx, by, bw, bottom - by, list(bullets), SIZE["body"], TEXT, line=LINE["body"], bullets=True,
                 gap_after=16)
        self._finish(s, notes, refs)

    def _cards_row(self, s, x, y, w, max_h, cards, page):
        """글 묶음 (소제목 + 불릿) 을 (x, y) 부터 폭 w 에 나란히 — 상자 없이. 돌려주는 것: 높이."""
        n, gap = len(cards), 48
        cw = (w - gap * (n - 1)) / max(n, 1)
        th = SIZE["card_title"] * LINE["card"]

        def body_h(card):
            body = card.get("body", [])
            lines = _lines(list(body), SIZE["card_body"], cw - SIZE["card_body"])
            return lines * SIZE["card_body"] * LINE["card"] + 10 * max(len(body) - 1, 0)

        hgt = th + 12 + max(body_h(c) for c in cards)
        if hgt > max_h:
            self._warn(page, "글 묶음이 칸을 넘음 — 줄이거나 노트로")
            hgt = max_h
        for i, card in enumerate(cards):
            cx = x + i * (cw + gap)
            _textbox(s, cx, y, cw, th, card["title"], SIZE["card_title"], INK, bold=True, line=LINE["card"])
            _textbox(s, cx, y + th + 12, cw, hgt - th - 12, list(card.get("body", [])),
                     SIZE["card_body"], TEXT, line=LINE["card"], bullets=True, gap_after=10)
        return hgt

    def cards(self, title, subtitle="", cards=(), key="", notes="", part=None, refs=()):
        """글 묶음 2–3 개 — cards: [{"title": 영어 소제목, "body": [불릿, …]}]. 상자는 없다.
        그림 없이 글만 있는 장이 되므로 경고한다 — 글 묶음은 equation·figure_bullets 와 함께 쓰는 쪽이 낫다."""
        s, top, bottom = self._start(title, subtitle, key, "cards", part=part)
        self._warn(self._page[s.slide_id], "그림 없는 장 — 글 묶음만 있으면 AI 가 만든 장처럼 보인다 (그림이나 표와 함께)")
        self._cards_row(s, MX, top, CW, bottom - top, cards, self._page[s.slide_id])
        self._finish(s, notes, refs)

    def equation(self, title, subtitle="", eq="", where="", image="", label="", caption="", cards=(), bullets=(),
                 key="", notes="", box=True, takeaway="", part=None, refs=()):
        """식 장 — 위에 전체 폭 식 (가운데·원래 크기, 번호 없음, 아래 where 한 문장).
        box=True 면 채움 없는 검정 2¼ pt 상자로 묶고, False 면 상자 없이 둔다.
        그 아래 왼쪽 그림 + 오른쪽 글 묶음 (또는 불릿). 그림이 없으면 글 묶음·불릿이 전체 폭을 쓴다.
        eq: tools/equations.py 로 만든 PNG 경로, 또는 LaTeX 글 (여러 줄은 리스트, 맞출 자리에 &) — 글이면 그 자리에서 만든다."""
        s, top, bottom = self._start(title, subtitle, key, "equation", part=part)
        page = self._page[s.slide_id]
        png = _equation_png(eq)
        with Image.open(png) as im:
            (iw, ih), dpi = im.size, (im.info.get("dpi") or (600, 600))[0]
        ew, eh = iw / dpi * 144, ih / dpi * 144            # 그림의 실제 크기 (in) → 캔버스 px (1 in = 144 px)
        pad = 36
        room = CW - 2 * pad
        if ew > room:
            self._warn(page, "식이 상자보다 넓음 — 줄을 나눔")
            ew, eh = room, eh * room / ew
        wn = _lines(where, SIZE["card_body"], room) if where else 0
        if wn > 2:
            self._warn(page, "where 가 2 줄을 넘음 — 줄이거나 노트로")
        wh = wn * SIZE["card_body"] * LINE["card"]
        box_h = 2 * pad + eh + (16 + wh if where else 0)
        if box:
            _rect(s, MX, top, CW, box_h, None, BOX_LINE, radius=8, width=BOX_W)
        _picture(s, png, MX + (CW - ew) / 2, top + pad, ew, eh, "equation")
        if where:
            _textbox(s, MX + pad, top + pad + eh + 16, room, wh, where, SIZE["card_body"], GREY, line=LINE["card"],
                     align=PP_ALIGN.CENTER)
        y = top + box_h + 40
        if (image or cards or bullets) and bottom - y < 120:
            self._warn(page, "식 아래 자리가 모자람 — 그림·글을 줄이거나 장을 나눔")
        elif image or cards or bullets:
            x, w = MX, CW
            if image:
                pw = self._figure(s, MX, y, CW * 0.45, bottom - y,
                                  {"image": image, "label": label, "takeaway": takeaway, "caption": caption}, page)
                x = MX + pw + 64
                w = MX + CW - x
            if cards:
                self._cards_row(s, x, y, w, bottom - y, cards, page)
            elif bullets:
                _textbox(s, x, y, w, bottom - y, list(bullets), SIZE["body"], TEXT, line=LINE["body"], bullets=True,
                         gap_after=16)
        self._finish(s, notes, refs)

    def table(self, title, subtitle="", header=(), rows=(), widths=None, key="", notes="", width=None,
              highlight=(), part=None, refs=()):
        """논문식 표 — header: 머리 줄, rows: 줄 목록, widths: 칸 폭 비율.
        이 연구가 나은 칸은 **굵게**, 꼭 짚을 값만 highlight=[(줄, 칸)] 으로 빨강 (줄은 rows 의 번호, 0 부터)."""
        s, top, bottom = self._start(title, subtitle, key, "table", part=part)
        _, hgt = _table(s, MX, top, width or CW, header, rows, widths, highlight=highlight)
        if top + hgt > bottom:
            self._warn(self._page[s.slide_id], "표가 칸을 넘음 — 행을 줄이거나 장을 나눔")
        self._finish(s, notes, refs)

    def summary(self, title, subtitle="", columns=(), closing="", notes="", part=None):
        """요약 — 필요할 때만 쓴다 (꼴은 내용에 맞게 — 표나 그림 장으로 맺어도 된다).
        얇은 구분선 아래 열 (소제목 + 불릿), 맨 아래 회색 마무리 문장."""
        s, top, bottom = self._start(title, subtitle, "", "summary", part=part)
        _rule(s, MX, top, MX + CW)
        n, gap = len(columns), 64
        cw = (CW - gap * (n - 1)) / max(n, 1)
        th = SIZE["card_title"] * LINE["card"]
        ch = SIZE["body"] * LINE["body"] if closing else 0
        for i, col in enumerate(columns):
            x = MX + i * (cw + gap)
            _textbox(s, x, top + 32, cw, th, col["title"], SIZE["card_title"], INK, bold=True, line=LINE["card"])
            _textbox(s, x, top + 32 + th + 12, cw, bottom - top - 32 - th - 12 - ch - 40, list(col.get("body", [])),
                     SIZE["body"], TEXT, line=LINE["body"], bullets=True, gap_after=12)
        if closing:
            _textbox(s, MX, bottom - ch, CW, ch, closing, SIZE["body"], MUTED, line=LINE["body"])
        self._finish(s, notes)

    def appendix_index(self, rows, title="Appendix", subtitle="Question Index", notes="",
                       header=("질문", "Appendix", "장")):
        """Appendix 첫 장 (쪽 번호 A0) — rows: [(질문, "A1"), …]. 장 번호 칸은 save() 때 채운다. 8 줄이 넘으면 표 둘."""
        self._appendix = 0
        s, top, bottom = self._start(title, subtitle, "", "appendix_index", page="A0")
        rows = list(rows)
        half = (len(rows) + 1) // 2
        groups = [rows] if len(rows) <= 8 else [rows[:half], rows[half:]]
        tw = (CW - 48) / 2
        for g, part in enumerate(groups):
            tbl, hgt = _table(s, MX + g * (tw + 48), top, tw, header, [(q, lab, "") for q, lab in part],
                              widths=(62, 22, 16))
            self._index_cells += [(tbl.cell(r, 2), lab) for r, (_, lab) in enumerate(part, start=1)]
            if top + hgt > bottom:
                self._warn("A0", "색인 표가 칸을 넘음")
        self._finish(s, notes)

    # 저장 ------------------------------------------------------------------
    def _check_deck(self):
        """덱 전체 — 빨간 줄이 너무 많은지, 같은 꼴이 여러 장 이어지는지."""
        most = len(self.parts) + 1 if self.parts else max(2, self._body // 4)
        if self._keys > most:
            self._warn("덱", f"빨간 줄 {self._keys} 장 — 파트를 맺는 장에만 ({most} 장 이하로)")
        run = []
        for no, kind in self._kinds + [(None, None)]:
            if run and kind == run[-1][1]:
                run.append((no, kind))
                continue
            if len(run) >= 4:
                self._warn(f"{run[0][0]}–{run[-1][0]}", f"같은 꼴 ({run[0][1]}) 이 {len(run)} 장 이어짐 — 장마다 내용에 맞는 꼴로")
            run = [(no, kind)]

    def save(self, path):
        for cell, lab in self._index_cells:
            no = self._labels.get(lab)
            if no is None:
                self._warn("A0", f"색인의 {lab} 장이 덱에 없음")
            cell.text_frame.paragraphs[0].runs[0].text = str(no if no is not None else "?")
        self._check_deck()
        cp = self.prs.core_properties                      # 작성자·회사 같은 개인 정보를 남기지 않는다
        cp.title = self.title
        cp.author = cp.last_modified_by = cp.subject = cp.keywords = cp.comments = cp.category = ""
        cp.revision = 1
        cp.created = cp.modified = datetime.datetime(2000, 1, 1)
        self.prs.save(str(path))
        _fixed_zip_dates(path)
        print(f"{Path(path).name}: {self._count} 장")
        for w in self.warnings:
            print("  경고 " + w)
        return self.warnings
