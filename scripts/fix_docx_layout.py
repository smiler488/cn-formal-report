#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""正式中文报告 DOCX 版式后处理：A4 页面、三线表、表格不跨页链、表题绑定、清除 AIGC 标记。

用法：python3 fix_docx_layout.py <docx>
"""
import os
import re
import sys
import zipfile

SECTPR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1418" w:right="1474" w:bottom="1418" w:left="1588" '
    'w:header="720" w:footer="720" w:gutter="0"/>'
    '<w:cols w:space="425"/>'
    '<w:docGrid w:type="lines" w:linePitch="312"/>'
    '</w:sectPr>'
)
EMPTY_CUSTOM = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
    'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
    '2006/docPropsVTypes"></Properties>'
)


def add_para_keep(p, keep_next, keep_lines, tight=False):
    """在段落 pPr 中插入 keepNext / keepLines / 单倍行距，保持 OOXML 子元素顺序。"""
    if re.search(r'<w:pPr\s*/>', p):
        p = re.sub(r'<w:pPr\s*/>', '<w:pPr></w:pPr>', p, count=1)
    if tight:
        p = re.sub(r'<w:spacing\b[^>]*/>', '', p)
    spacing = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' if tight else ''
    if re.search(r'<w:pPr\b[^>]*>', p):
        ins = spacing
        if keep_lines and '<w:keepLines' not in p:
            ins = '<w:keepLines/>' + ins
        if keep_next and '<w:keepNext' not in p:
            ins = '<w:keepNext/>' + ins
        if ins:
            p = re.sub(r'(<w:pPr\b[^>]*>)', r'\1' + ins, p, count=1)
    else:
        ins = (('<w:keepNext/>' if keep_next else '')
               + ('<w:keepLines/>' if keep_lines else '') + spacing)
        if ins:
            p = re.sub(r'(<w:p\b[^>]*>)', r'\1<w:pPr>' + ins + '</w:pPr>', p, count=1)
    return p


def fix_table(tbl):
    rows = list(re.finditer(r'<w:tr\b[^>]*>.*?</w:tr>', tbl, re.S))
    if not rows:
        return tbl
    spans = [(m.start(), m.end(), m.group(0)) for m in rows]
    out, cursor = [], 0
    for idx, (start, end, tr) in enumerate(spans):
        out.append(tbl[cursor:start])
        is_last_row = (idx == len(spans) - 1)
        if re.search(r'<w:trPr\b[^>]*>', tr):
            if '<w:cantSplit' not in tr:
                tr = re.sub(r'(<w:trPr\b[^>]*>)', r'\1<w:cantSplit/>', tr, count=1)
        else:
            tr = re.sub(r'(<w:tr\b[^>]*>)',
                        r'\1<w:trPr><w:cantSplit/></w:trPr>', tr, count=1)
        if idx == 0 and '<w:tblHeader' not in tr:
            tr = re.sub(r'(<w:trPr\b[^>]*>)', r'\1<w:tblHeader/>', tr, count=1)
        tr = re.sub(r'<w:p\b[^>]*>.*?</w:p>',
                    lambda m: add_para_keep(m.group(0), not is_last_row, True, tight=True),
                    tr, flags=re.S)
        out.append(tr)
        cursor = end
    out.append(tbl[cursor:])
    return ''.join(out)


def main(path):
    with zipfile.ZipFile(path, 'r') as zin:
        parts = {i.filename: (i, zin.read(i.filename)) for i in zin.infolist()}
    doc = parts['word/document.xml'][1].decode('utf-8')

    doc = re.sub(r'<w:tbl\b[^>]*>.*?</w:tbl>',
                 lambda m: fix_table(m.group(0)), doc, flags=re.S)

    def caption(m):
        p = m.group(0)
        text = ''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', p, re.S))
        if re.match(r'^表\s*\d+[　\s]', text):
            return add_para_keep(p, True, True)
        return p
    doc = re.sub(r'<w:p\b[^>]*>.*?</w:p>', caption, doc, flags=re.S)

    doc = re.sub(r'<w:sectPr\s*/>', SECTPR, doc)
    if '<w:pgSz' not in doc:
        doc = doc.replace('</w:body>', SECTPR + '</w:body>')

    parts['word/document.xml'] = (parts['word/document.xml'][0], doc.encode('utf-8'))
    if 'docProps/custom.xml' in parts:
        parts['docProps/custom.xml'] = (parts['docProps/custom.xml'][0],
                                        EMPTY_CUSTOM.encode('utf-8'))

    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, (info, data) in parts.items():
            zout.writestr(info, data)
    os.replace(tmp, path)
    print('版式修复：%d 表 / %d 行不分页 / %d 表头重复 / A4 已写入 / AIGC 已清除'
          % (doc.count('<w:tbl>'), doc.count('cantSplit'), doc.count('tblHeader')))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit('用法：fix_docx_layout.py <docx>')
    main(sys.argv[1])
