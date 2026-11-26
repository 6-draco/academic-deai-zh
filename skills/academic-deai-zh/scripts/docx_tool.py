#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""academic-deai-zh 的 .docx 辅助脚本（依赖 python-docx）。

子命令：
  extract <in.docx> <paras.json>
      抽取文档主体段落（不含表格单元格），输出 JSON 列表，每项含：
        index, style, text（纯文本）, marked（加粗片段以 **…** 标出，仅混合加粗段落有）,
        chars, skip, skip_reason, complex, complex_reason
      自动标记 skip：空段、标题（样式或“一、/（一）/1.”编号、整段加粗短段）、居中短段（题目/署名）、
      作者信息行、关键词行、目录、图表题注、参考文献标题及其后列表。

  check <edits.json>
      对每条改写做硬性核查：锚点（引注、年份/数字、英文术语、[n] 编号）是否全部保留，
      引号内术语是否保留（缺失记为 WARN），字数比是否在 85%–115%。有 ERROR 时返回码 5。

  apply <in.docx> <edits.json> <out.docx> [--force]
      edits.json：[{"index": 12, "original": "原段落纯文本", "revised": "改写文本，可用 **加粗**"}]
      - original 必须与该段当前纯文本完全一致，否则中止（防错位）。
      - 段落样式、缩进、行距不变；普通文字沿用原段首个非加粗 run 的字体字号，
        **…** 片段沿用原段加粗 run 的格式（原段无加粗时在普通格式上加粗）。
      - 含超链接/脚注/域代码/公式/图片，或存在加粗以外的格式差异（上标、斜体、字号等）的段落
        视为 complex，默认拒绝写回；--force 强制按上述规则重建。
      - 输出路径不得与输入相同。
"""
import copy
import json
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    import docx  # python-docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
except ImportError:
    sys.stderr.write("缺少 python-docx，请先执行：pip install python-docx\n")
    sys.exit(2)

CAPTION_RE = re.compile(r"^\s*(表|图|Table|Figure|Fig\.)\s*[0-9一二三四五六七八九十]+([.\-－–][0-9]+)?(\s|$)")
KEYWORD_RE = re.compile(r"^\s*(关键词|关键字|Key\s*words?)\s*[:：]", re.I)
META_RE = re.compile(r"^\s*(作者简介|通信地址|通讯地址|联系电话|联系方式|电子邮箱|邮箱|基金项目|收稿日期|作者单位|中图分类号|文献标识码|文章编号)")
REF_HEAD_RE = re.compile(r"^\s*(参考文献|References|Bibliography)\s*$", re.I)
HEADING_NUM_RE = re.compile(r"^\s*(第[一二三四五六七八九十百]+[章节]|[一二三四五六七八九十]+[、.．]|[（(][一二三四五六七八九十]+[）)]|\d+(\.\d+)*[.．、]?\s)")
COMPLEX_TAGS = {
    "w:hyperlink": "超链接",
    "w:footnoteReference": "脚注引用",
    "w:endnoteReference": "尾注引用",
    "w:fldChar": "域代码",
    "w:fldSimple": "域代码",
    "m:oMath": "公式",
    "w:drawing": "图片/图形",
}


def _fmt_sig(run, with_bold=True):
    f = run.font
    color = None
    try:
        if f.color is not None and f.color.type is not None:
            color = str(f.color.rgb)
    except Exception:
        pass
    sig = (f.name, f.size, f.italic, f.underline, f.superscript, f.subscript, color)
    return sig + ((bool(f.bold),) if with_bold else ())


def _text_runs(par):
    return [r for r in par.runs if r.text.strip()]


def complex_reason(par):
    xml = par._p.xml
    reasons = [label for tag, label in COMPLEX_TAGS.items() if "<" + tag in xml]
    if len({_fmt_sig(r, with_bold=False) for r in _text_runs(par)}) > 1:
        reasons.append("存在加粗以外的格式差异（上标/斜体/字号/字体等）")
    return "；".join(dict.fromkeys(reasons))


def mixed_bold(par):
    return len({bool(r.font.bold) for r in _text_runs(par)}) > 1


def marked_text(par):
    out = []
    for r in par.runs:
        if r.font.bold and r.text.strip():
            out.append("**" + r.text + "**")
        else:
            out.append(r.text)
    return re.sub(r"\*\*\*\*", "", "".join(out))


def classify(par, text, in_refs):
    style = par.style.name if par.style is not None else ""
    sl = style.lower()
    runs = _text_runs(par)
    all_bold = bool(runs) and all(r.font.bold for r in runs)
    short = len(text.strip()) <= 50
    ends_sentence = text.rstrip().endswith(("。", "；", "！", "？", ".", "”"))
    if not text.strip():
        return "空段"
    if REF_HEAD_RE.match(text):
        return "参考文献标题"
    if in_refs:
        return "参考文献列表"
    if sl.startswith("heading") or "标题" in style or sl == "title":
        return "标题（样式）"
    if sl.startswith("toc") or "目录" in style:
        return "目录"
    if "caption" in sl or "题注" in style or CAPTION_RE.match(text):
        return "图表题注"
    if KEYWORD_RE.match(text):
        return "关键词行"
    if META_RE.match(text):
        return "作者/文章信息"
    if short and par.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return "题目/署名（居中短段）"
    if short and HEADING_NUM_RE.match(text) and not ends_sentence:
        return "标题（编号）"
    if short and all_bold:
        return "标题（整段加粗）"
    return ""


def extract(src, out):
    d = docx.Document(src)
    items, in_refs = [], False
    for i, p in enumerate(d.paragraphs):
        text = p.text
        why = classify(p, text, in_refs)
        if why == "参考文献标题":
            in_refs = True
        cr = complex_reason(p)
        item = {
            "index": i,
            "style": p.style.name if p.style is not None else "",
            "text": text,
            "chars": count_chars(text),
            "skip": bool(why),
            "skip_reason": why,
            "complex": bool(cr),
            "complex_reason": cr,
        }
        if mixed_bold(p):
            item["marked"] = marked_text(p)
        items.append(item)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(items, fh, ensure_ascii=False, indent=2)
    body = [x for x in items if not x["skip"]]
    print("共 %d 段；建议处理正文 %d 段（含加粗强调 %d 段，complex %d 段）-> %s"
          % (len(items), len(body), sum("marked" in x for x in body),
             sum(x["complex"] for x in body), out))
    for x in body:
        print("  [%d] %d字%s%s  %s…" % (x["index"], x["chars"],
                                       " 加粗" if "marked" in x else "",
                                       " COMPLEX" if x["complex"] else "", x["text"][:24]))


# ---------------- check ----------------

def strip_marks(s):
    return s.replace("**", "")


def count_chars(s):
    return len(re.sub(r"\s", "", strip_marks(s)))


ANCHOR_PATTERNS = [
    ("引注", re.compile(r"[\u4e00-\u9fa5A-Za-z·、]{1,20}[（(]\d{4}[a-z]?[）)]")),
    ("引注", re.compile(r"[（(][^（）()]{0,40}?\d{4}[^（）()]{0,40}?[）)]")),
    ("编号", re.compile(r"\[\d+(?:\s*[-–—,，]\s*\d+)*\]")),
    ("图表", re.compile(r"[表图]\s*\d+(?:[.\-]\d+)?")),
    ("英文", re.compile(r"[A-Za-z][A-Za-z\-']*(?:\s+[A-Za-z][A-Za-z\-']*)*")),
    ("数字", re.compile(r"\d+(?:\.\d+)?%?")),
]
QUOTE_RE = re.compile(r"“[^“”]{1,40}”")


def anchors(text):
    t = strip_marks(text)
    found = []
    for kind, pat in ANCHOR_PATTERNS:
        for m in pat.findall(t):
            m = m if isinstance(m, str) else m[0]
            if (kind, m) not in found:
                found.append((kind, m))
    return found


def check_one(orig, rev):
    errs, warns = [], []
    o, r = strip_marks(orig), strip_marks(rev)
    for kind, a in anchors(orig):
        need = o.count(a)
        have = r.count(a)
        if have < need:
            errs.append("%s缺失：%s（原 %d 处，现 %d 处）" % (kind, a, need, have))
    for q in dict.fromkeys(QUOTE_RE.findall(o)):
        if q not in r:
            warns.append("引号术语未原样保留：" + q)
    new_nums = set(re.findall(r"\d+(?:\.\d+)?", r)) - set(re.findall(r"\d+(?:\.\d+)?", o))
    if new_nums:
        errs.append("出现原文没有的数字：" + "、".join(sorted(new_nums)))
    for bad in ("！", "!"):
        if bad in r and bad not in o:
            errs.append("新增感叹号")
    co, cr = count_chars(orig), count_chars(rev)
    ratio = cr / co if co else 0
    if not 0.85 <= ratio <= 1.15:
        errs.append("字数比 %.1f%% 超出 85%%–115%%" % (ratio * 100))
    return co, cr, ratio, errs, warns


def check(edits_path):
    with open(edits_path, encoding="utf-8") as fh:
        edits = json.load(fh)
    n_err = 0
    for e in edits:
        co, cr, ratio, errs, warns = check_one(e["original"], e["revised"])
        status = "ERROR" if errs else ("WARN" if warns else "OK")
        n_err += bool(errs)
        print("[%d] %s  原文 %d 字 → 改后 %d 字（%.1f%%）" % (e["index"], status, co, cr, ratio * 100))
        for m in errs:
            print("    ERROR " + m)
        for m in warns:
            print("    WARN  " + m)
    print("合计 %d 段，ERROR %d 段" % (len(edits), n_err))
    return 5 if n_err else 0


# ---------------- apply ----------------

def _rpr_of(run):
    rpr = run._r.rPr
    return copy.deepcopy(rpr) if rpr is not None else None


def _set_bold(rpr_el, on):
    if rpr_el is None:
        return
    for tag in ("w:b", "w:bCs"):
        for el in rpr_el.findall(qn(tag)):
            rpr_el.remove(el)
    if on:
        rpr_el.append(rpr_el.makeelement(qn("w:b"), {}))
        rpr_el.append(rpr_el.makeelement(qn("w:bCs"), {}))


def rebuild_runs(par, revised):
    runs = par.runs
    text_runs = _text_runs(par) or runs
    plain_src = next((r for r in text_runs if not r.font.bold), text_runs[0] if text_runs else None)
    bold_src = next((r for r in text_runs if r.font.bold), None)
    plain_rpr = _rpr_of(plain_src) if plain_src is not None else None
    if plain_rpr is not None and bold_src is None:
        _set_bold(plain_rpr, False)
    if bold_src is not None:
        bold_rpr = _rpr_of(bold_src)
    else:
        bold_rpr = copy.deepcopy(plain_rpr) if plain_rpr is not None else None
        _set_bold(bold_rpr, True)
    for r in runs:
        r._r.getparent().remove(r._r)
    parts = revised.split("**")
    if len(parts) % 2 == 0:
        raise ValueError("段落 **加粗** 标记不成对")
    for k, seg in enumerate(parts):
        if not seg:
            continue
        new = par.add_run(seg)
        rpr = bold_rpr if k % 2 == 1 else plain_rpr
        if rpr is not None:
            old = new._r.rPr
            if old is not None:
                new._r.remove(old)
            new._r.insert(0, copy.deepcopy(rpr))
        elif k % 2 == 1:
            new.bold = True


def apply(src, edits_path, out, force=False):
    if os.path.abspath(src).lower() == os.path.abspath(out).lower():
        sys.stderr.write("拒绝覆盖原文件：输出路径必须与输入不同\n")
        sys.exit(3)
    with open(edits_path, encoding="utf-8") as fh:
        edits = json.load(fh)
    d = docx.Document(src)
    paras = d.paragraphs
    errors, refused = [], []
    for e in edits:
        idx = int(e["index"])
        if idx < 0 or idx >= len(paras):
            errors.append("段落 %d 越界" % idx)
            continue
        p = paras[idx]
        if p.text != strip_marks(e["original"]):
            errors.append("段落 %d 原文校验失败（文本不一致）" % idx)
            continue
        if e["revised"].count("**") % 2:
            errors.append("段落 %d 的 **加粗** 标记不成对" % idx)
        cr = complex_reason(p)
        if cr and not force:
            refused.append("段落 %d: %s" % (idx, cr))
    if errors or refused:
        for m in errors:
            sys.stderr.write("[错误] " + m + "\n")
        for m in refused:
            sys.stderr.write("[complex，未加 --force] " + m + "\n")
        sys.stderr.write("未写出任何文件。\n")
        sys.exit(4)
    for e in edits:
        rebuild_runs(paras[int(e["index"])], e["revised"])
    d.save(out)
    print("已写入 %d 段 -> %s" % (len(edits), out))


def main(argv):
    if len(argv) >= 3 and argv[0] == "extract":
        extract(argv[1], argv[2])
    elif len(argv) >= 2 and argv[0] == "check":
        sys.exit(check(argv[1]))
    elif len(argv) >= 4 and argv[0] == "apply":
        apply(argv[1], argv[2], argv[3], force="--force" in argv[4:])
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv[1:])
