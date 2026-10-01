# 党建月度文档生成 Skill

一个用于 Codex 的党建工作文档生成技能。它根据月度重点工作清单和往月记录，生成党支部大会、集中学习、主题党日、谈心谈话等中文 DOCX 材料，并尽量保留本单位原有的 Word 格式。

本项目重点解决三件事：

- 从 Excel 工作清单确定本月活动、主题、完成时间和合并关系。
- 从往月同类型 DOCX 学习标题层级、正文结构、篇幅和 Word 样式。
- 通过脚本继承历史模板生成新文档，并在交付前逐页进行视觉检查。

> 仓库只包含通用工作流和脚本，不包含任何真实党建文件、党员信息、单位名称或活动照片。

## 功能

- 支持支部党员大会或支委会、集中学习、主题党日、谈心谈话、支部书记讲党课等材料。
- 支持“第三次集中学习和谈心谈话”等跨活动合并文件。
- 自动识别清单中的 `xx党支部`、`xx` 等占位符。
- 复制最近同类型 DOCX 作为模板，避免重新定义页面和样式。
- 提供清单、历史文档结构和字数检查脚本。
- 从月度清单生成活动顺序、规范文件名和模板候选的交付计划。
- 提供 DOCX 页面、样式、字体、字距、行距和段距画像脚本。
- 提供基于 JSON 规格的模板继承生成脚本。
- 把模板残留、月份错误、命名错误和逐页渲染检查纳入质量门。

## 安装

将仓库克隆到 Codex 个人技能目录：

```powershell
$skillRoot = if ($env:CODEX_HOME) {
    Join-Path $env:CODEX_HOME "skills"
} else {
    Join-Path $env:USERPROFILE ".codex\skills"
}

git clone https://github.com/z825572/party-building-monthly-docs.git `
    (Join-Path $skillRoot "party-building-monthly-docs")
```

重新打开 Codex 任务后，可以通过 `$party-building-monthly-docs` 显式调用。

## 使用

示例请求：

```text
使用 $party-building-monthly-docs。
根据本月重点工作清单和往月学习记录，生成 <党支部名称> 本月党建 DOCX 材料。
第三次集中学习和谈心谈话合并为一个文件，并沿用往月 Word 格式。
```

技能会先读取工作清单和最近三个月的同类记录，再生成文档和进行逐页检查。

## 脚本

### 检查清单和历史记录

```bash
python scripts/inspect_monthly_inputs.py \
  --root <党建材料根目录> \
  --month <月份筛选> \
  --details
```

脚本会输出：

- 工作清单工作表和单元格内容。
- 历史 DOCX 的非空段落数、字符数、表格数、页面尺寸和页边距。
- 段落样式、首行缩进、中文字体和字号等格式信息。

### 从 JSON 规格生成 DOCX

准备类似以下的 `spec.json`：

```json
{
  "output_dir": "outputs",
  "documents": [
    {
      "template": "references/previous-record.docx",
      "filename": "党支部9月第一次集中学习.docx",
      "paragraphs": [
        {
          "style": "党建标题",
          "text": "学习内容记录："
        },
        {
          "style": "党建正文",
          "text": "全体党员集中学习了……",
          "line_spacing_pt": 31
        },
        {
          "style": "党建正文",
          "blank": true
        },
        {
          "style": "党建正文",
          "text": "讨论发言摘要：",
          "bold": true,
          "first_line_indent_pt": 0
        }
      ]
    }
  ]
}
```

执行：

```bash
python scripts/build_docx_from_spec.py --spec spec.json
```

可用参数：

- `--output-dir`：覆盖 JSON 中的输出目录。
- `--overwrite`：允许覆盖同名文件。
- `--dry-run`：只校验规格，不写入 DOCX。

段落支持 `style`、`text`、`blank`、`bold`、`italic`、`alignment`、`first_line_indent_pt`、`left_indent_pt`、`right_indent_pt`、`space_before_pt`、`space_after_pt`、`line_spacing_pt`、`page_break_before`、`keep_with_next` 和 `keep_together`。默认应让模板样式控制格式，只对确实需要差异的段落写直接格式。

### 生成月度交付计划

```bash
python scripts/plan_monthly_docs.py \
  --root <党建材料根目录> \
  --month 9月 \
  --branch-name <党支部名称> \
  --lookback 3
```

计划会按清单行输出：

- 活动顺序、类别、完成时间和清单行号。
- 补齐党支部名称后的 DOCX 文件名。
- 最近 3 个月中同类别、同次序的文件模板候选。
- 无法从清单确定的 `xxxx` 主题和缺失模板警告。

计划结果只用于生成前确认，最终输出目录不应保留计划 JSON。

### 验收交付目录

```bash
python scripts/validate_delivery.py \
  --directory <交付目录> \
  --month 9月 \
  --branch-name <党支部名称> \
  --expected-count 6
```

验收器会检查文件数量、月份、党支部名称、`xxxx` 占位符、DOCX ZIP 完整性、文件能否重新打开和正文是否为空。它不替代逐页渲染检查。

如果计划器输出了 JSON，可再传 `--plan <plan.json>`，验收器会同时核对缺失文件和计划外文件。

### 剖析 DOCX 格式

```bash
python scripts/profile_docx_formatting.py \
  --root <党建材料根目录> \
  --month 6月 \
  --month 7月 \
  --month 8月 \
  --include-source \
  --json
```

脚本会输出：

- 页面尺寸、方向、页边距、页眉页脚距离和文档网格。
- 使用到的段落样式及其继承关系。
- 有效字体、字号、加粗、颜色和字符间距。
- 行距、段前段后、首行缩进和空段统计。
- 按正式记录和学习原文分组的格式聚合结果。

2026 年 6 至 8 月的实测汇总见 `references/word-format-profile-2026-06-08.md`。

## 依赖

- Python 3.11+
- `openpyxl`
- `python-docx`

```bash
python -m pip install -r requirements.txt
```

## 目录结构

```text
party-building-monthly-docs/
|-- SKILL.md
|-- agents/
|   `-- openai.yaml
|-- references/
|   |-- writing-patterns.md
|   `-- word-format-profile-2026-06-08.md
|-- scripts/
|   |-- build_docx_from_spec.py
|   |-- plan_monthly_docs.py
|   |-- profile_docx_formatting.py
|   |-- validate_delivery.py
|   `-- inspect_monthly_inputs.py
|-- requirements.txt
`-- LICENSE
```

## 写作与隐私

- 仓库中的篇幅范围和段落示例只作为默认参考，最近一期同类文件优先。
- 不应提交真实党员姓名、单位内部材料、活动照片、会议记录或未脱敏的工作清单。
- `.gitignore` 默认忽略 DOCX、XLSX、PDF 和常见图片格式，以降低误提交风险。
- 使用公开仓库前，请再次人工检查内容是否符合本单位的信息公开要求。

## 验证

检查脚本语法：

```bash
python -m compileall scripts
```

使用 `skill-creator` 的校验器检查技能结构：

```bash
python <skill-creator>/scripts/quick_validate.py .
```

## License

[MIT](LICENSE)
