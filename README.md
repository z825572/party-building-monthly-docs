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

## 使用说明

### 1. 准备材料

建议按月份整理党建材料。调用 Skill 时提供材料根目录、目标月份和党支部名称即可，不需要手工转换 Word 或 Excel。

```text
党建材料根目录/
|-- 9月/
|   |-- 9月份重点工作清单.xlsx
|   |-- 9月学习原文/
|   `-- 其他当月资料.docx
|-- 8月/
|   |-- 8月党支部大会.docx
|   |-- 8月第一次集中学习.docx
|   `-- ...
`-- 7月/
    `-- ...
```

为了提高生成质量，目录中应尽量包含：

- 当月重点工作清单，其中保留活动主题、活动内容、完成时间和党支部占位符。
- 当月需要学习或引用的原文材料、通知和领导讲话。
- 最近一至三个月已经审核通过的同类 DOCX 记录。
- 需要写入文档的准确党支部名称和其他固定信息。

### 2. 在 Codex 中调用

安装并重新打开 Codex 任务后，在请求中显式写出 `$party-building-monthly-docs`，并提供以下关键信息：

```text
使用 $party-building-monthly-docs。
材料根目录：<党建材料根目录>
目标月份：<月份>
党支部名称：<党支部名称>
重点工作清单：<清单路径或文件名>
历史参考范围：<例如最近 3 个月>
```

如果合并规则、文件名或活动形式有特殊要求，也应直接写在请求中。例如“第三次集中学习和谈心谈话合并为一个文件”或“主题党日结合本月节日和近期学习内容开展”。

### 3. 技能执行过程

调用后，技能会按照以下顺序工作：

1. 读取重点工作清单，确认活动数量、类别、顺序、主题、完成时间和合并要求。
2. 读取最近同类型 DOCX，分析文章结构、篇幅、字体字号、字符间距、行距、段距和页面样式。
3. 生成月度交付计划，列出每份文件的名称、模板候选和需要人工确认的问题。
4. 基于最近同类型 Word 模板生成 DOCX，并将清单占位符替换为实际党支部名称。
5. 检查 XML 内容、字段名称、ZIP 完整性、重新打开能力和正文是否为空。
6. 将全部页面渲染为图片，逐页检查错页、截断、空标题、模板残留和格式异常。

技能最终交付的是 DOCX 材料，不交付 JSON、分析表、构建脚本、PDF 或页面图片等中间文件。正式提交前仍应由使用者核对活动日期、参与人员、政策表述和本单位的信息公开要求。

## 推荐提示词

以下提示词可以按实际情况修改路径、月份、党支部名称和活动数量后直接使用。

### 完整生成本月材料

```text
使用 $party-building-monthly-docs。

请根据以下材料完成本月党建工作记录：
- 材料根目录：<党建材料根目录>
- 目标月份：<月份>
- 党支部名称：<党支部名称>
- 重点工作清单：<清单路径或文件名>
- 历史参考：最近 3 个月同类型 DOCX

要求：
1. 先阅读重点工作清单和往月记录，确认每种材料的写作方法、文章结构、字数范围和 Word 格式。
2. 按照清单生成党支部大会、集中学习、主题党日和谈心谈话等全部文档。
3. 原文材料中的内容优先于历史记录，历史记录只用于参考结构和格式，不机械复制旧内容。
4. 第三次集中学习和谈心谈话合并为一个 DOCX 文件。
5. 主题党日结合当月节日或纪念日、近期学习内容和本单位业务场景开展。
6. 将所有占位符替换为指定党支部名称，输出规范命名的 DOCX。
7. 生成后执行交付校验，并把所有页面渲染后逐页检查。
```

### 先生成写作与格式分析

```text
使用 $party-building-monthly-docs。

暂时不要生成文档。请先分析 <材料根目录> 中 <月份范围> 的党建记录，并按材料类型总结：
1. 每种材料的标题层级、正文结构、常用段落和惯用表述。
2. 各类文档的典型字数范围以及最短、最长记录。
3. 页面尺寸、页边距、字体、字号、字符间距、行距、段前段后和首行缩进。
4. 党支部大会、集中学习、主题党日和谈心谈话之间的格式差异。
5. 生成新文档时应优先继承哪些模板，以及有哪些容易出现的模板残留。

分析完成后给出结论，等我确认后再生成 DOCX。
```

### 只生成月度交付计划

```text
使用 $party-building-monthly-docs。

请先为 <月份> 生成交付计划，不要创建 DOCX。
材料根目录：<党建材料根目录>
党支部名称：<党支部名称>
模板参考范围：最近 3 个月

计划中请列出：
- 清单原始顺序、活动类别、活动主题、完成时间和清单行号。
- 每份文件的规范名称，以及需要合并到同一文件的活动。
- 推荐使用的历史模板及推荐理由。
- 清单中的 xxxx 主题、缺失模板、命名冲突等需要确认的问题。
- 按合并规则计算后的最终 DOCX 数量。
```

### 补发或重新生成部分文档

```text
使用 $party-building-monthly-docs。

请只处理 <需要补发或修改的文档名称>，不要重新生成已经确认无误的其他文件。
材料根目录：<党建材料根目录>
目标月份：<月份>
党支部名称：<党支部名称>
需要修改的内容：<具体说明>

请沿用最近同类型 DOCX 的模板、字体、字号、行距和段距，修改后重新进行交付校验和逐页渲染检查。
```

### 开展特色主题党日

```text
使用 $party-building-monthly-docs。

请生成 <月份> 主题党日记录。
党支部名称：<党支部名称>
本月节日或纪念日：<节日、纪念日或活动主题>
近期学习内容：<学习主题或原文材料>
业务结合点：<本单位近期重点工作或服务场景>

活动内容应体现节日背景、理论学习、交流实践和活动成效，避免写成普通集中学习，也不要编造未提供的人物事迹或活动细节。格式沿用最近一次主题党日 DOCX。
```

### 验收已有交付目录

```text
使用 $party-building-monthly-docs。

请验收 <交付目录> 中为 <党支部名称> 生成的 <月份> 党建 DOCX。
要求：
1. 对照重点工作清单检查是否有遗漏、重复或多余文件。
2. 检查文件名、月份、党支部名称和 xxxx 占位符。
3. 检查每个 DOCX 的 ZIP 完整性、能否重新打开、正文是否为空。
4. 将所有页面渲染后逐页检查错页、截断、空标题和模板残留。
5. 发现问题时直接修复，并重新执行完整校验。
```

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
