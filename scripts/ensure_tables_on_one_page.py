#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""渲染实测驱动的「表格整表同页」修复。

流程：DOCX --Aspose(Word 布局)--> PDF --PyMuPDF--> 每张表各行所在页
     -> 跨页表在其表题段落前插入分页符 -> 重渲染复测，直至全部整表同页。
用法：python3 ensure_tables_on_one_page.py <docx> [--max-pass 5]
依赖：pip install aspose-words pymupdf
"""
import html
import os
import re
import sys
import zipfile

try:
    import aspose.words as aw
    import pymupdf
except ImportError as exc:
    sys.exit('缺少依赖：pip install aspose-words pymupdf （%s）' % exc)

BREAK_PARA = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
              'w:lineRule="auto"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>')


def read_doc(path):
    with zipfile.ZipFile(path, 'r') as zin:
        return {i.filename: (i, zin.read(i.filename)) for i in zin.infolist()}


def write_doc(path, parts):
    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, (info, data) in parts.items():
            zout.writestr(info, data)
    os.replace(tmp, path)


def xml_text(fragment):
    return html.unescape(''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>',
                                           fragment, re.S)))


def table_row_keys(doc):
    out = []
    for ti, tbl in enumerate(re.findall(r'<w:tbl\b[^>]*>.*?</w:tbl>', doc, re.S)):
        keys = []
        for tr in re.findall(r'<w:tr\b[^>]*>.*?</w:tr>', tbl, re.S):
            cells = re.findall(r'<w:tc\b[^>]*>.*?</w:tc>', tr, re.S)
            keys.append(re.sub(r'\s+', '', ''.join(xml_text(c) for c in cells))[:18])
        out.append((ti, keys))
    return out


def render_pages(docx_path, pdf_path):
    doc = aw.Document(docx_path)
    doc.save(pdf_path, aw.SaveFormat.PDF)
    pdf = pymupdf.open(pdf_path)
    pages = [re.sub(r'\s+', '', pg.get_text()) for pg in pdf]
    pdf.close()
    return pages


def check(docx_path):
    parts = read_doc(docx_path)
    doc = parts['word/document.xml'][1].decode('utf-8')
    pdf_path = docx_path + '.layout.pdf'
    try:
        pages = render_pages(docx_path, pdf_path)
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
    result = {}
    for ti, keys in table_row_keys(doc):
        hits = set()
        for key in keys:
            if not key:
                continue
            for pi, page_text in enumerate(pages):
                if key in page_text:
                    hits.add(pi + 1)
        result[ti] = (sorted(hits), len(keys))
    return result


def insert_breaks(docx_path, targets):
    parts = read_doc(docx_path)
    doc = parts['word/document.xml'][1].decode('utf-8')
    positions = [m.start() for m in
                 re.finditer(r'<w:tbl\b[^>]*>.*?</w:tbl>', doc, re.S)]
    inserts = []
    for ti in targets:
        if ti >= len(positions):
            continue
        before = doc[:positions[ti]]
        paras = list(re.finditer(r'<w:p\b[^>]*>.*?</w:p>', before, re.S))
        anchor = positions[ti]
        if paras and re.match(r'^表\s*\d+[　\s]', xml_text(paras[-1].group(0))):
            anchor = paras[-1].start()
        inserts.append(anchor)

    changed = False
    for pos in sorted(set(inserts), reverse=True):
        if '<w:br w:type="page"/>' in doc[max(0, pos - 260):pos]:
            continue
        doc = doc[:pos] + BREAK_PARA + doc[pos:]
        changed = True
    if changed:
        parts['word/document.xml'] = (parts['word/document.xml'][0], doc.encode('utf-8'))
        write_doc(docx_path, parts)
    return changed


def main():
    if len(sys.argv) < 2:
        sys.exit('用法：ensure_tables_on_one_page.py <docx> [--max-pass 5]')
    path = sys.argv[1]
    max_pass = 5
    if '--max-pass' in sys.argv:
        max_pass = int(sys.argv[sys.argv.index('--max-pass') + 1])

    for pass_no in range(1, max_pass + 1):
        result = check(path)
        print('第%d轮：' % pass_no,
              {'表%d' % (ti + 1): pgs for ti, (pgs, _) in result.items()})
        bad = [ti for ti, (pgs, _) in result.items() if len(pgs) > 1]
        if not bad:
            print('全部表格整表同页')
            return 0
        print('  跨页表：', ['表%d' % (ti + 1) for ti in bad], '→ 插入前置分页符')
        if not insert_breaks(path, bad):
            print('分页符已存在仍跨页，需按内容拆分为多个短表')
            return 2
    print('超过最大轮次仍存在跨页表')
    return 1


if __name__ == '__main__':
    sys.exit(main())
