# cn-formal-report

面向邓良超（中国农业大学生物学博士后，作物模型与农业人工智能方向）的中文正式报告写作与 Word 成品交付技能。

## 解决什么问题

把反复被指出的规范固化成可执行检查，避免每次写报告都重新犯同样的错：

- 主标题必须单行，不得把「调研报告」等词拆行；
- 表格必须整表同页，且以 Word 布局渲染实测验收，而不是只检查 OOXML 属性；
- 公文层次「一、→（一）→ 1.」，标题用名词性短语、禁判断句与情态词；
- 正文零辅助性文字，只交付可直接提交的成品；
- 每次用户批评都转成规则 + 自动检查 + 版本递增 + GitHub 同步。

## 使用

```bash
# 依赖
pip install aspose-words pymupdf

# 一键生成并校验
python3 ~/.workbuddy/skills/cn-formal-report/scripts/make_docx.py 报告.md 报告.docx

# 单独校验已有文档
python3 ~/.workbuddy/skills/cn-formal-report/scripts/verify_report.py 报告.docx --md 报告.md

# 同步规则变更
~/.workbuddy/skills/cn-formal-report/scripts/sync_to_github.sh "fix: 简述"
```

技能描述覆盖「写报告／改报告／公文／调研报告／Word 成品／标题分行／表格跨页／太 AI」等触发词，命中后自动加载。规则细节在 `references/`，脚本在 `scripts/`。

## 目录

```
SKILL.md                     触发说明、工作流、硬规则速查、自升级协议
references/01-format-rules.md    版式硬规范
references/02-heading-and-wording.md  结构、标题、用词
references/03-deliverable.md     交付形态与自检清单
references/04-user-profile.md    身份与场景
references/05-feedback-rules.md  反馈台账（每次批评追加）
scripts/make_docx.py             一键生成 + 校验
scripts/build_report.py          Markdown → 正式 HTML
scripts/fix_docx_layout.py       A4 / 三线表 / 不跨页属性 / 清除 AIGC
scripts/ensure_tables_on_one_page.py  渲染实测驱动的跨页修复
scripts/verify_report.py         交付前全量校验
scripts/sync_to_github.sh        git 提交推送
CHANGELOG.md / VERSION          版本与升级记录
```

## 升级协议

收到格式或用词批评时：先修当前交付件 → 写入反馈台账 → 在 `verify_report.py` 增加检查项并回归 → 更新 CHANGELOG 与 VERSION → 执行同步脚本推送。详见 `SKILL.md` 第四节。

GitHub：https://github.com/smiler488/cn-formal-report
