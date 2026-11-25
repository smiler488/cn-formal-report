#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""正式中文报告交付前全量校验（结构 + 用词 + 渲染实测）。

用法：python3 verify_report.py <docx> [--md <input.md>]
失败返回非零退出码；渲染依赖 aspose-words、pymupdf，缺失直接报错不静默跳过。
"""
import html
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

try:
    import aspose.words as aw
    import pymupdf
except ImportError as exc:
    sys.exit('缺少渲染依赖：pip install aspose-words pymupdf （%s）' % exc)

BANNED_TEXT = [
    '用途：', '核验说明', '核验等级', '编制日期', '本报告', '需人工核实',
    '应避免的表述', '一页纸会谈备忘', '留档材料', '清稿提醒', '结尾免责声明',
    '一、结论', '结论摘要', '核心判断', '总体判断', '综上所述', '总的来说',
    '值得注意的是', '赋能', '抓手', '闭环', '颠覆', '需要的话我可以',
]
BANNED_HEADING_WORDS = ['已', '须', '应', '需', '将', '为', '有',
                        '判断', '结论', '摘要', '核心判断']
LEGIT_COMPOUNDS = ['有关', '有效', '有机', '有序', '应用', '对应',
                   '具有', '为主', '所为', '认为']


def xml_text(fragment):
    return html.unescape(''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>',
                                           fragment, re.S)))


def norm(s):
    return re.sub(r'\s+', '', s)


def strip_quoted(s):
    """剔除引号内内容：官方文件原名、章节标题中的用词不受自撰用词规则约束。"""
    return re.sub(r'“[^”]*”|《[^》]*》|「[^」]*」|『[^』]*』', '', s)


def row_keys(tbl):
    keys = []
    for tr in re.findall(r'<w:tr\b[^>]*>.*?</w:tr>', tbl, re.S):
        cells = re.findall(r'<w:tc\b[^>]*>.*?</w:tc>', tr, re.S)
        keys.append(norm(''.join(xml_text(c) for c in cells))[:18])
    return keys


def check_structure(path):
    results = []
    z = zipfile.ZipFile(path)
    raw = z.read('word/document.xml').decode('utf-8')
    try:
        ET.fromstring(raw)
        results.append(('XML 合法', True, ''))
    except Exception as exc:
        results.append(('XML 合法', False, str(exc)))

    results.append(('A4 与页边距', 'w:w="11906"' in raw and 'pgMar' in raw,
                    '不得留空 sectPr'))

    tbls = re.findall(r'<w:tbl\b[^>]*>.*?</w:tbl>', raw, re.S)
    bad_rows = []
    for i, t in enumerate(tbls, 1):
        rows = re.findall(r'<w:tr\b[^>]*>.*?</w:tr>', t, re.S)
        for j, r in enumerate(rows, 1):
            pr = re.search(r'<w:trPr\b[^>]*>(.*?)</w:trPr>', r, re.S)
            if not pr or 'cantSplit' not in pr.group(1):
                bad_rows.append('表%d 第%d行 cantSplit 缺失或不在 trPr 内' % (i, j))
        pr0 = re.search(r'<w:trPr\b[^>]*>(.*?)</w:trPr>', rows[0], re.S)
        if not pr0 or 'tblHeader' not in pr0.group(1):
            bad_rows.append('表%d 表头未设重复' % i)
        if 'keepNext' in rows[-1]:
            bad_rows.append('表%d 末行未断链' % i)
    results.append(('表格同页结构属性', not bad_rows, '；'.join(bad_rows)))

    custom = z.read('docProps/custom.xml').decode('utf-8') if 'docProps/custom.xml' in z.namelist() else ''
    results.append(('无 AIGC 元数据', 'AIGC' not in custom, ''))

    text = xml_text(raw)
    prose = strip_quoted(text)
    hits = [w for w in BANNED_TEXT if w in prose]
    results.append(('无辅助性文字与套话', not hits, '、'.join(hits)))

    colors = (re.findall(r'<w:color w:val="([0-9A-Fa-f]{6})"', raw)
              + re.findall(r'<w:fill="([0-9A-Fa-f]{6})"', raw))
    bad_colors = [c for c in colors if c.upper() not in ('000000', 'FFFFFF')]
    highlights = [v for v in re.findall(r'<w:highlight w:val="(\w+)"', raw)
                  if v.lower() not in ('none', 'auto')]
    results.append(('三线表与正文无彩色', not bad_colors and not highlights,
                    '色彩值 %s / 高亮 %s' % (bad_colors[:4], highlights[:4])))

    # 正文洁净度：Word 视图辅助符号不是文档缺陷，但生成链路不得写入真实杂字符。
    paras = re.findall(r'<w:p\b[^>]*>.*?</w:p>', raw, re.S)
    ptexts = [xml_text(q) for q in paras]
    issues = []
    n_empty = sum(1 for t in ptexts if not t.strip())
    n_solo = sum(1 for t in ptexts
                 if t.strip() and re.fullmatch(r'[\s.。·•\-—]*', t.strip()))
    n_head = sum(1 for t in ptexts if t[:1] in (' ', '.', '·', '。', '\u3000'))
    n_tail = sum(1 for t in ptexts if t[-1:] in (' ', '.', '·', '\u3000'))
    n_hidden = raw.count('w:vanish')
    n_br = raw.count('<w:br')
    n_shape = (raw.count('<w:drawing') + raw.count('<w:pict')
               + raw.count('v:shape'))
    for label, n in [('空段落', n_empty), ('孤立标点段落', n_solo),
                     ('段首空格点号', n_head), ('段尾空格点号', n_tail),
                     ('隐藏文字', n_hidden), ('手动换行', n_br),
                     ('形状文本框', n_shape)]:
        if n:
            issues.append('%s %d' % (label, n))
    results.append(('正文无杂字符与隐藏对象', not issues, '；'.join(issues)))
    return results, raw, tbls, text


def check_markdown(md_path):
    results = []
    if not md_path or not os.path.exists(md_path):
        return results
    src = open(md_path, encoding='utf-8').read()
    heads = [l.strip() for l in src.split('\n') if re.match(r'^#{2,4}\s+', l.strip())]
    bad = []
    for h in heads:
        core = re.sub(r'^#{2,4}\s+', '', h)
        core = re.sub(r'^[一二三四五六七八九十]+、|^（[一二三四五六七八九十]+）|^\d+\.\s*', '', core)
        if re.search(r'[，。；：、！？]', core):
            bad.append('%s（含标点）' % core)
            continue
        clean = core
        for w in LEGIT_COMPOUNDS:
            clean = clean.replace(w, '')
        hit = [w for w in BANNED_HEADING_WORDS if w in clean]
        if hit:
            bad.append('%s（含 %s）' % (core, '、'.join(hit)))
    results.append(('公文标题规范', not bad, '；'.join(bad[:6])))

    lines = src.split('\n')
    heads = []
    for idx, line in enumerate(lines):
        m = re.match(r'^(#{2,4})\s+', line.strip())
        if m:
            heads.append((idx, len(m.group(1)), line.strip()))
    empty = []
    for i, (line_no, level, title) in enumerate(heads):
        end = len(lines)
        for j in range(i + 1, len(heads)):
            if heads[j][1] <= level:
                end = heads[j][0]
                break
        body = '\n'.join(lines[line_no + 1:end]).strip()
        if len(body) < 40:
            empty.append(re.sub(r'^#+\s+', '', title)[:20])
    results.append(('无空节', not empty, '、'.join(empty)))
    return results


def check_render(path, raw, tbls, md_path):
    results = []
    pdf_path = path + '.verify.pdf'
    try:
        doc = aw.Document(path)
        doc.save(pdf_path, aw.SaveFormat.PDF)
        pdf = pymupdf.open(pdf_path)
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

    # 主标题单行
    if md_path and os.path.exists(md_path):
        m = re.search(r'^#\s+(.+)$', open(md_path, encoding='utf-8').read(), re.M)
        title = m.group(1).strip() if m else ''
    else:
        first_p = re.search(r'<w:p\b[^>]*>.*?</w:p>', raw, re.S)
        title = xml_text(first_p.group(0)) if first_p else ''
    lines = []
    for block in pdf[0].get_text('dict')['blocks']:
        if 'lines' in block:
            for line in block['lines']:
                lines.append(norm(''.join(s['text'] for s in line['spans'])))
    title_norm = norm(title)
    ok = bool(title_norm) and any(title_norm in ln for ln in lines)
    results.append(('主标题单行', ok, '标题：%s' % title))

    # 表格整表同页
    pages = [norm(pg.get_text()) for pg in pdf]
    detail, all_ok = [], True
    for i, t in enumerate(tbls, 1):
        pgs = set()
        for key in row_keys(t):
            if not key:
                continue
            for pi, pt in enumerate(pages):
                if key in pt:
                    pgs.add(pi + 1)
        detail.append('表%d→%s' % (i, sorted(pgs)))
        if len(pgs) != 1:
            all_ok = False
    results.append(('全部表格整表同页（渲染实测）', all_ok, '；'.join(detail)))
    results.append(('无空白页', all(len(pg) > 40 for pg in pages), '页数 %d' % pdf.page_count))
    pdf.close()
    return results


def main():
    if len(sys.argv) < 2:
        sys.exit('用法：verify_report.py <docx> [--md <input.md>]')
    path = sys.argv[1]
    md_path = sys.argv[sys.argv.index('--md') + 1] if '--md' in sys.argv else None

    struct, raw, tbls, text = check_structure(path)
    md = check_markdown(md_path)
    render = check_render(path, raw, tbls, md_path)
    all_results = struct + md + render

    print('=== 正式报告交付校验 ===')
    failed = 0
    for name, ok, detail in all_results:
        print('%s %s%s' % ('通过' if ok else '失败', name,
                           ('　' + detail) if (detail and not ok) else ''))
        if not ok:
            failed += 1
    print('正文 %d 字 / 表格 %d 张 / 超链接 %d 条'
          % (len(text), len(tbls), raw.count('HYPERLINK')))
    print('结论：%s' % ('全部通过' if not failed else '%d 项未通过，不得交付' % failed))
    return 0 if not failed else 1


if __name__ == '__main__':
    sys.exit(main())
