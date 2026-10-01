# 党建月度文档生成 Skill

一个用于 Codex 的中文党建工作文档生成技能。

它只需要一份月度重点工作清单，就可以生成党支部大会、支委会、集中学习、主题党日、谈心谈话、支部书记讲党课等整套 DOCX 材料，并统一使用内置的标准 Word 模板。

> 往月 DOCX 不再是生成时的必需输入。技能已经把现有记录中提炼出的页面、样式、字体、行距和段距固化为脱敏标准模板，以后每月只需要提供清单。

## 主要特点

- 单文件输入：默认只需要一份重点工作清单 `.xlsx`。
- 不再依赖往月文档：不会搜索、复制或要求提供历史记录。
- 内置统一模板：所有文档共享相同的页面、标题、正文和段落规范。
- 自动读取清单：识别活动类别、主题、内容、参与人员、完成时间和合并关系。
- 自动规范文件名：补齐月份和党支部名称，删除 `xxxx` 等占位符。
- 支持合并记录：可将“第三次集中学习和谈心谈话”等活动合并为一个 DOCX。
- 支持姓名名单：清单附加列提供三个以上姓名时，讨论发言优先使用这些姓名。
- 自动结构验收：检查文件数量、命名、月份、党支部名称、ZIP 完整性和正文。
- 隐私友好：真实党支部名称保存在本机忽略配置中，不进入公开仓库。

## 工作效果

```text
输入：2026年9月重点工作清单.xlsx
  |
  |-- 支委会/支部党员大会
  |-- 第一次集中学习
  |-- 第二次集中学习
  |-- 主题党日
  |-- 谈心谈话和第三次集中学习（合并）
  |-- 第四次集中学习
  `-- 支部书记讲党课
  |
输出：7 份统一格式的 DOCX
```

如果清单中没有明确要求合并，则每个活动生成一个 DOCX；如果清单明确要求合并，则按一个文件生成。

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

安装 `python-docx` 和 `openpyxl`：

```bash
python -m pip install -r requirements.txt
```

重新打开 Codex 任务后，通过 `$party-building-monthly-docs` 调用。

## 首次配置

如果清单本身已经包含真实党支部名称，技能可以直接识别。

如果清单使用 `xx党支部` 占位符，只需首次运行一次配置命令，把真实党支部名称保存在本机：

```bash
python scripts/generate_monthly_docs.py \
  --checklist "<重点工作清单.xlsx>" \
  --branch-name "<党支部名称>" \
  --save-config
```

该名称保存在技能目录下的 `party-building.local.json`。文件被 `.gitignore` 忽略，不会提交到 GitHub。

配置格式可参考公开的通用示例 `party-building.local.example.json`。

配置完成之后，后续月份只需要提供清单文件。

## 快速开始

只提供一份重点工作清单：

```bash
python scripts/generate_monthly_docs.py \
  --checklist "<2026年9月重点工作清单.xlsx>"
```

脚本会自动完成：

1. 从文件名或清单内容推断年份、月份。
2. 读取所有活动类别、主题、内容和完成时间。
3. 自动识别党支部名称、合并活动和输出文件名。
4. 从 `templates/standard-party-template.docx` 创建统一格式的 DOCX。
5. 先生成学习记录、主题党日、谈心谈话和讲党课正文。
6. 执行文件数量、命名、月份、占位符和 DOCX 完整性检查。

默认输出目录为清单同级的：

```text
<年份>年<月份>月党建生成/
```

例如：

```text
2026年9月党建生成/
|-- <党支部名称>9月份支部党员大会或支委会.docx
|-- <党支部名称>9月第一次集中学习.docx
|-- <党支部名称>9月第二次集中学习.docx
|-- <党支部名称>9月主题党日.docx
|-- <党支部名称>9月谈心谈话和9月第三次集中学习.docx
|-- <党支部名称>9月第四次集中学习.docx
`-- <党支部名称>9月支部书记讲党课.docx
```

示例中的 `<党支部名称>` 在运行时替换为本机配置或用户提供的名称。

指定输出目录：

```bash
python scripts/generate_monthly_docs.py \
  --checklist "<重点工作清单.xlsx>" \
  --output-dir "<输出目录>"
```

覆盖已有文件：

```bash
python scripts/generate_monthly_docs.py \
  --checklist "<重点工作清单.xlsx>" \
  --overwrite
```

只验证解析和规格，不写出 DOCX：

```bash
python scripts/generate_monthly_docs.py \
  --checklist "<重点工作清单.xlsx>" \
  --dry-run
```

## 使用说明

### 1. 准备清单

清单可以沿用原来的 Excel 结构，至少应包含：

| 列名 | 作用 |
| --- | --- |
| 活动类别 | 判断生成支委会、集中学习、主题党日还是谈心谈话 |
| 活动主题 | 确定文件名、活动名称和占位符 |
| 活动内容 | 提取学习主题、课程名称或谈话主题 |
| 参与人员 | 写入学习记录或活动元数据 |
| 完成时间 | 确定活动日期或月份 |
| 要求/备注 | 识别合并、录入和其他特别要求 |

清单附加列如果提供多个规范姓名，讨论发言会优先采用这些姓名；没有名单时使用通用身份，不虚构真实人物。

### 2. 在 Codex 中调用

只需要给出清单路径即可：

```text
使用 $party-building-monthly-docs。
根据这份重点工作清单生成全部党建 DOCX：
<重点工作清单.xlsx>
```

技能会读取清单、使用统一模板生成文档、执行结构校验，并在交付前完成逐页渲染检查。

### 3. 查看生成结果

生成完成后，优先检查：

- 文件数量是否与清单活动及合并要求一致。
- 文件名中的月份、党支部名称和活动次序是否正确。
- 党员姓名、活动日期、参与人员和专有材料名称是否来自清单。
- 清单没有提供的讲师、群众反馈、荣誉和案例是否被错误补写。
- 每个页面是否有标题截断、孤立标题、空白页或错误换行。

### 4. 正文边界

清单只提供课程名或专题名时，生成器会在该主题范围内写出完整的学习记录、讨论摘要和谈话内容，但不会虚构：

- 未提供的讲师、时长、具体互动和课堂案例。
- 未提供的人物事迹、荣誉、奖惩或重大事件。
- 未提供的群众评价、家庭隐私、投诉或纠纷。
- 超出清单和当地政策要求的具体承诺。

如果用户后续补充学习原文或活动资料，可以只重新生成对应文档。

## 推荐提示词

### 只提供清单，生成整月材料

```text
使用 $party-building-monthly-docs。

只根据这份重点工作清单生成本月全部党建 DOCX：
<重点工作清单.xlsx>

要求：
1. 不要求提供往月文档，不查找历史模板。
2. 使用技能内置的标准 Word 模板统一生成全部文件。
3. 自动处理活动类别、合并关系、文件名、月份和党支部名称。
4. 生成后执行结构校验，并对全部页面进行渲染检查。
```

### 指定输出目录并覆盖旧版本

```text
使用 $party-building-monthly-docs。

根据 <重点工作清单.xlsx> 重新生成全部党建 DOCX。
输出目录：<输出目录>
如果目录已有同名文件，请覆盖并重新校验。
```

### 首次设置党支部名称

```text
使用 $party-building-monthly-docs。

请使用 <重点工作清单.xlsx> 完成首次配置。
本机党支部名称：<党支部名称>
保存配置后生成全部材料。
以后我只提供重点工作清单，不再重复提供党支部名称和往月文档。
```

### 清单内容更新后重新生成

```text
使用 $party-building-monthly-docs。

我已经更新了 <重点工作清单.xlsx>，请重新读取整份清单并生成全部 DOCX。
不要沿用上一次输出的正文；只保留本机党支部名称配置和统一 Word 模板。
```

### 先生成交付计划，不写 DOCX

```text
使用 $party-building-monthly-docs。

请分析 <重点工作清单.xlsx>，先列出：
1. 活动数量、类别、次序和合并关系。
2. 每份 DOCX 的规范文件名。
3. 主题党日主题、谈心谈话对象和讲党课主题。
4. 按合并规则计算后的最终文件数量。

暂时不要创建 DOCX，等我看完计划后再生成。
```

### 验收已有交付目录

```text
使用 $party-building-monthly-docs。

请验收 <交付目录> 中的党建 DOCX。
对照 <重点工作清单.xlsx> 检查文件数量、命名、月份、党支部名称、占位符、ZIP 完整性和正文。
发现问题时直接修复，并重新执行结构校验和逐页渲染检查。
```

## 统一模板

标准模板位于：

```text
templates/standard-party-template.docx
```

模板只包含脱敏的页面和样式定义，不包含真实单位、姓名、会议内容或活动照片。

主要样式：

| 样式 | 用途 | 主要格式 |
| --- | --- | --- |
| `党建标题` | 长记录一级标题 | 方正小标宋简体 22 pt，居中，固定行距 36 pt |
| `党建标题1` | 栏名和短标题 | 宋体 22 pt，加粗，居中，固定行距 36 pt |
| `党建正文` | 常规正文 | 仿宋_GB2312 16 pt，首行缩进 32 pt，固定行距 31 pt |
| `全文一级大标题` | 长记录大标题 | 黑体 22 pt，加粗，居中，固定行距 32 pt |
| `正文一级标题` | 长记录一级标题 | 黑体 18 pt，加粗，固定行距 28 pt |
| `正文二级标题` | 长记录二级标题 | 黑体 17 pt，加粗，固定行距 28 pt |
| `正文 文本` | 长记录正文 | 宋体 16 pt，首行缩进 32 pt，固定行距 28 pt |

需要统一修改整套格式时，编辑：

```text
scripts/build_standard_template.py
```

然后重新生成模板：

```bash
python scripts/build_standard_template.py --overwrite
```

## 脚本说明

### 单清单完整生成

```bash
python scripts/generate_monthly_docs.py --checklist <重点工作清单.xlsx>
```

这是常规使用的唯一入口。

### 生成标准模板

```bash
python scripts/build_standard_template.py --overwrite
```

用于维护模板；正常生成文档时不需要运行。

### 校验交付目录

```bash
python scripts/validate_delivery.py \
  --directory <交付目录> \
  --month <月份> \
  --branch-name <党支部名称> \
  --expected-count <文件数>
```

### JSON 规格生成

需要对少量段落做特殊排版时，仍可使用：

```bash
python scripts/build_docx_from_spec.py --spec <spec.json>
```

这是高级入口，不是月度生成的必经步骤。

### 格式与输入诊断

以下脚本只用于维护模板或排查问题，不属于用户生成流程：

- `scripts/inspect_monthly_inputs.py`
- `scripts/profile_docx_formatting.py`
- `scripts/plan_monthly_docs.py`

它们可以读取历史文件来分析格式，但正常生成不会调用这些历史文件。

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
|   |-- build_standard_template.py
|   |-- generate_monthly_docs.py
|   |-- inspect_monthly_inputs.py
|   |-- plan_monthly_docs.py
|   |-- profile_docx_formatting.py
|   `-- validate_delivery.py
|-- templates/
|   |-- README.md
|   `-- standard-party-template.docx
|-- party-building.local.example.json
|-- requirements.txt
`-- LICENSE
```

## 隐私

- 公开仓库不包含真实党员姓名、单位内部材料、会议记录、活动照片或原始清单。
- `party-building.local.json` 被忽略，不会提交真实党支部名称。
- `.gitignore` 默认忽略 DOCX、XLSX、PDF 和常见图片。
- 使用公开技能和各平台模型前，请再次核对材料是否符合本单位的信息公开要求。

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
