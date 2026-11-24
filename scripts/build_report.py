#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把正式中文报告 Markdown 转为排版 HTML（供 html-to-doc 转换）。

用法：python3 build_report.py <input.md> <output.html>
内置规则：黑字、宋体正文、黑体标题、三线表、主标题按字数自动定字号保证单行。
"""
import html
import math
import re
import sys

BODY_FONT = "font-family:'SimSun','宋体',serif; font-size:12pt; color:#000000; line-height:1.65;"
H_FONT = "font-family:'SimHei','黑体',sans-serif; color:#000000;"
USABLE_WIDTH_PT = 442.0  # A4 11906tw - 左1588tw - 右1474tw


def title_font_pt(text):
    """按标题宽度自动定字号，保证单行；上限 18pt，下限 14pt。"""
    em = 0.0
    for ch in text:
        em += 0.5 if ord(ch) < 0x2E80 else 1.0
    em += 1.0  # 中西文自动间距
    size = min(16, int(USABLE_WIDTH_PT / (em * 1.02)))
    if size < 14:
        print('警告：标题过长，需精简文字：%s' % text, file=sys.stderr)
        size = 14
    return size


def inline(text):
    t = html.escape(text, quote=False)
    t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
    t = re.sub(r'(?<!\*)\*([^*\n]+?)\*(?!\*)', r'<em>\1</em>', t)
    t = re.sub(r'(https?://[^\s|<]+)', r'<a href="\1" style="color:#000000;">\1</a>', t)
    return t


def is_sep(row):
    cells = [c.strip() for c in row.strip().strip('|').split('|')]
    return bool(cells) and all(re.fullmatch(r':?-{2,}:?', c or '') for c in cells)


def split_row(row):
    return [c.strip() for c in row.strip().strip('|').split('|')]


def build(src, dst):
    lines = open(src, encoding='utf-8').read().split('\n')
    first_title = next((l.strip()[2:].strip() for l in lines
                        if re.match(r'^#\s+', l.strip())), '')
    h1_size = title_font_pt(first_title) if first_title else 16

    out, i, n = [], 0, len(lines)
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue

        if s == '---':
            out.append('<p style="%s text-align:center; margin:10pt 0;">—　—　—</p>' % BODY_FONT)
            i += 1
            continue

        m = re.match(r'^(#{1,4})\s+(.*)$', s)
        if m:
            lvl, txt = len(m.group(1)), inline(m.group(2))
            if lvl == 1:
                st = H_FONT + ('font-size:%dpt; text-align:center; margin:16pt 0 8pt; '
                               'white-space:nowrap;' % h1_size)
            elif lvl == 2:
                st = H_FONT + 'font-size:14pt; text-align:left; margin:16pt 0 8pt;'
            elif lvl == 3:
                st = H_FONT + 'font-size:12.5pt; margin:12pt 0 6pt;'
            else:
                st = H_FONT + 'font-size:12pt; margin:10pt 0 4pt;'
            out.append('<h%d style="%s">%s</h%d>' % (lvl, st, txt, lvl))
            i += 1
            continue

        if s.startswith('|'):
            block = []
            while i < n and lines[i].strip().startswith('|'):
                block.append(lines[i].strip())
                i += 1
            header, body_rows = None, []
            for idx, row in enumerate(block):
                if is_sep(row):
                    continue
                if idx == 0:
                    header = split_row(row)
                else:
                    body_rows.append(split_row(row))
            out.append('<table style="width:100%%; border-collapse:collapse; '
                       'border-top:1pt solid #000000; border-bottom:1pt solid #000000; '
                       'margin:4pt 0 8pt; page-break-inside:avoid; line-height:1.05;">')
            if header:
                out.append('<thead><tr>')
                for c in header:
                    out.append('<th style="%s font-size:8pt; line-height:1.05; '
                               'font-weight:bold; text-align:left; padding:1pt 3pt; '
                               'border-bottom:0.75pt solid #000000; word-break:break-all;">%s</th>'
                               % (BODY_FONT, inline(c)))
                out.append('</tr></thead>')
            out.append('<tbody>')
            for cells in body_rows:
                out.append('<tr>')
                for c in cells:
                    out.append('<td style="%s font-size:8pt; line-height:1.05; padding:1pt 3pt; '
                               'vertical-align:top; word-break:break-all;">%s</td>'
                               % (BODY_FONT, inline(c)))
                out.append('</tr>')
            out.append('</tbody></table>')
            continue

        if re.match(r'^-\s+', s) or re.match(r'^\d+\.\s+', s):
            ordered = bool(re.match(r'^\d+\.\s+', s))
            tag, items = ('ol' if ordered else 'ul'), []
            while i < n:
                cur = lines[i].strip()
                if ordered and re.match(r'^\d+\.\s+', cur):
                    items.append(re.sub(r'^\d+\.\s+', '', cur))
                elif not ordered and re.match(r'^-\s+', cur):
                    items.append(re.sub(r'^-\s+', '', cur))
                else:
                    break
                i += 1
            out.append('<%s style="%s margin:6pt 0 8pt; padding-left:2em;">' % (tag, BODY_FONT))
            for it in items:
                out.append('<li style="%s margin:3pt 0; text-indent:0;">%s</li>'
                           % (BODY_FONT, inline(it)))
            out.append('</%s>' % tag)
            continue

        para = [s]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (not nxt or nxt == '---' or nxt.startswith('|') or nxt.startswith('#')
                    or re.match(r'^-\s+', nxt) or re.match(r'^\d+\.\s+', nxt)):
                break
            para.append(nxt)
            i += 1
        text = ' '.join(para)
        if re.match(r'^表\s*\d+[　\s]', text):
            out.append('<p style="%s font-size:10.5pt; margin:10pt 0 4pt; '
                       'text-indent:0;">%s</p>' % (H_FONT, inline(text)))
        elif re.match(r'^〔[^〕]+〕', text):
            out.append('<p style="%s font-size:10.5pt; line-height:1.4; margin:3pt 0; '
                       'text-indent:0; padding-left:1.5em;">%s</p>' % (BODY_FONT, inline(text)))
        elif text.startswith('*') and text.endswith('*') and len(text) > 2 and '**' not in text:
            out.append('<p style="%s font-size:10.5pt; margin:12pt 0 4pt; '
                       'text-indent:0;"><em>%s</em></p>' % (BODY_FONT, inline(text[1:-1])))
        else:
            out.append('<p style="%s text-indent:2em; margin:5pt 0;">%s</p>'
                       % (BODY_FONT, inline(text)))

    doc = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>正式报告</title>
<style>
@page { size: A4; margin: 25mm 26mm 25mm 28mm; }
body { font-family:'SimSun','宋体',serif; font-size:12pt; color:#000000; line-height:1.65; }
h1,h2,h3,h4 { font-family:'SimHei','黑体',sans-serif; color:#000000; }
table { border-top:1pt solid #000000; border-bottom:1pt solid #000000; }
th { border-bottom:0.75pt solid #000000; }
a { color:#000000; }
</style></head><body>
%s
</body></html>""" % '\n'.join(out)
    open(dst, 'w', encoding='utf-8').write(doc)
    print('HTML 已生成：%s（主标题 %dpt，%d 个块）' % (dst, h1_size, len(out)))


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit('用法：build_report.py <input.md> <output.html>')
    build(sys.argv[1], sys.argv[2])
