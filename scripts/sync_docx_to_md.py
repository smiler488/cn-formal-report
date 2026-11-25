#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把用户手工修订后的 Word 定稿反向同步为 Markdown 中间稿。

用法：python3 sync_docx_to_md.py <final.docx> <output.md>

规则：以 Word 定稿为唯一事实来源，保留逐段文字、句内加粗、表格、表题、
著录式引用目录与全部链接文本；不重新创作、不改写用词。生成链路
（make_docx.py）由此源稿重建时不会回退到旧写法。

背景：用户在 Word 里直接改成品后，.md 源稿即失效。禁止用旧源稿覆盖成品。
"""
import html
import os
import re
import sys
import zipfile


def xml_text(fragment):
    return html.unescape(''.join(re.findall(
        r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', fragment, re.S)))


def is_bold_run(run):
    pr = re.search(r'<w:rPr\b[^>]*>.*?</w:rPr>', run, re.S)
    if not pr:
        return False
    return bool(re.search(r'<w:b(?:\s+w:val="(?:1|true|on)")?\s*/>', pr.group(0)))


def para_markdown(p, keep_bold=True):
    """提取段落文本，粗体 run 转 **…**；相邻粗体 run 合并。"""
    parts = []
    for run in re.findall(r'<w:r\b[^>]*>.*?</w:r>', p, re.S):
        text = xml_text(run)
        if not text:
            continue
        bold = keep_bold and is_bold_run(run)
        if bold:
            if parts and parts[-1][1]:
                parts[-1] = (parts[-1][0] + text, True)
            else:
                parts.append((text, True))
        else:
            parts.append((text, False))
    out = ''.join(('**%s**' % t) if b else t for t, b in parts)
    return out.strip()


def heading_level(text, is_first):
    if is_first:
        return 1
    if re.match(r'^[一二三四五六七八九十]+、', text):
        return 2
    if re.match(r'^（[一二三四五六七八九十]+）', text):
        return 3
    return 0


def table_markdown(tbl):
    lines = []
    rows = re.findall(r'<w:tr\b[^>]*>.*?</w:tr>', tbl, re.S)
    width = 0
    for tr in rows:
        cells = re.findall(r'<w:tc\b[^>]*>.*?</w:tc>', tr, re.S)
        vals = [para_markdown(c).replace('|', r'\|') for c in cells]
        width = max(width, len(vals))
        lines.append('| ' + ' | '.join(vals) + ' |')
        if len(lines) == 1:
            lines.append('| ' + ' | '.join(['---'] * len(vals)) + ' |')
    return lines, len(rows), width


def convert(docx_path, md_path):
    z = zipfile.ZipFile(docx_path)
    raw = z.read('word/document.xml').decode('utf-8')
    body = raw[raw.find('<w:body>'):]
    items = re.findall(r'<w:tbl\b[^>]*>.*?</w:tbl>|<w:p\b[^>]*>.*?</w:p>', body, re.S)
    out, first_done, n_tables, n_refs = [], False, 0, 0
    for item in items:
        if item.startswith('<w:tbl'):
            block, _, _ = table_markdown(item)
            out.append('\n'.join(block))
            n_tables += 1
            continue
        plain = xml_text(item).strip()
        if not plain:
            continue
        lvl = heading_level(plain, not first_done)
        if lvl:
            out.append('%s %s' % ('#' * lvl, plain))
            first_done = True
            continue
        text = para_markdown(item)
        if not text:
            continue
        if re.match(r'^\[\d+[a-z]?\]', plain):
            n_refs += 1
        out.append(text)
    doc = '\n\n'.join(out) + '\n'
    open(md_path, 'w', encoding='utf-8').write(doc)
    print('反向同步完成：%s' % md_path)
    print('段落/表格块 %d 个 | 表格 %d 张 | 著录条目 %d 条 | 字符 %d'
          % (len(out), n_tables, n_refs, len(doc)))
    print('主标题：%s' % (out[0][2:] if out and out[0].startswith('# ') else '未识别'))
    return doc


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit('用法：sync_docx_to_md.py <final.docx> <output.md>')
    convert(sys.argv[1], sys.argv[2])
