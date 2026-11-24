#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键生成正式中文报告 Word 成品并校验。

用法：python3 make_docx.py <input.md> <output.docx>
链路：排版 HTML → DOCX → 版式修复 → 表格跨页实测修复 → 全量校验。
"""
import glob
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def run(args, label):
    print('\n--- %s ---' % label)
    proc = subprocess.run(args, text=True, capture_output=True)
    if proc.stdout:
        print(proc.stdout.strip())
    if proc.returncode != 0:
        print(proc.stderr.strip(), file=sys.stderr)
        sys.exit('%s 失败（退出码 %d）' % (label, proc.returncode))


def find_converter():
    patterns = [
        os.path.expanduser('~/.workbuddy/plugins/cache/**/html-to-doc/scripts/convert_doc.py'),
        '/Applications/WorkBuddy.app/Contents/Resources/app.asar.unpacked/resources/plugins/'
        'workbuddy-builtin/**/html-to-doc/scripts/convert_doc.py',
    ]
    hits = []
    for pat in patterns:
        hits.extend(glob.glob(pat, recursive=True))
    if not hits:
        sys.exit('未找到 WorkBuddy html-to-doc 的 convert_doc.py，请先启用 workbuddy-docx 插件')
    return sorted(hits)[-1]


def main():
    if len(sys.argv) != 3:
        sys.exit('用法：make_docx.py <input.md> <output.docx>')
    src, dst = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
    if not os.path.exists(src):
        sys.exit('输入文件不存在：%s' % src)

    html_path = os.path.join(tempfile.gettempdir(),
                             os.path.splitext(os.path.basename(src))[0] + '_formal.html')
    run([sys.executable, os.path.join(HERE, 'build_report.py'), src, html_path],
        '生成正式排版 HTML')
    run([sys.executable, find_converter(), '--html', html_path, '--output', dst,
         '--no-preflight', '--no-verify'], '转换为 DOCX')
    run([sys.executable, os.path.join(HERE, 'fix_docx_layout.py'), dst],
        'A4 / 三线表 / 不跨页属性 / 清除 AIGC')
    run([sys.executable, os.path.join(HERE, 'ensure_tables_on_one_page.py'), dst],
        '表格整表同页（渲染实测）')
    run([sys.executable, os.path.join(HERE, 'verify_report.py'), dst, '--md', src],
        '交付前全量校验')
    print('\n成品：%s' % dst)


if __name__ == '__main__':
    main()
