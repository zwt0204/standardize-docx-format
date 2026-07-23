# standardize-docx-format

面向 Codex 的 Microsoft Word `.docx` 格式标准化 Skill，重点解决毕业论文、学位论文、课程论文、技术报告等文档的复杂排版、模板套用和自动验收问题。

它既可以根据学校提供的官方 Word 模板工作，也可以把自然语言、PDF、网页或截图中的格式要求转换为结构化 JSON Profile，再确定性地应用到现有 DOCX 文件中。

> 这个项目首先是一个 Codex Skill，不是 Word 插件，也不是带图形界面的排版软件。核心脚本只依赖 Python 标准库，可以独立完成 DOCX 包审计、Profile 应用和结构验收；复杂需求识别、模板区域映射和未决规则处理由 Codex 工作流负责。

## 主要能力

- 保留原始 DOCX，默认输出新的 `.standardized.docx` 文件，拒绝原地覆盖。
- 保留图片、表格、书签、字段、关系、页眉页脚和未建模 OOXML 扩展。
- 支持 A3、A4、A5、Letter 等页面尺寸、横竖方向和逐节页边距。
- 支持封面、摘要、目录、正文、参考文献、致谢、附录等不同分节规则。
- 支持中文和西文字体分槽设置：
  - 中文宋体、黑体、楷体等写入 `w:eastAsia`。
  - 英文 Times New Roman 等写入 `w:ascii`、`w:hAnsi` 和 `w:cs`。
- 支持字号、粗体、斜体、颜色、段前段后、行距、首行缩进、悬挂缩进、两端对齐和分页控制。
- 支持罗马页码、阿拉伯页码和正文重新从 1 编号。
- 支持首页、普通页、偶数页分别设置页眉页脚。
- 支持中文论文多级自动编号：
  - `一、`
  - `（一）`
  - `1.`
  - `（1）`
  - `①`
- 支持自动目录、PAGE、NUMPAGES、SEQ、REF、PAGEREF、STYLEREF、CITATION、BIBLIOGRAPHY 等 Word 字段。
- 支持书签、题注、图表编号和交叉引用基础设施。
- 支持跨多个 run、页眉、页脚和文本框的精确占位符替换。
- 支持按文本、正则、现有样式、节、段落位置、表格或文本框位置映射语义样式。
- 支持删除模板教学说明、清理手写标题编号和统一正文格式。
- 支持结构、必备章节、字段、书签、页码、页眉页脚、占位符和中英文摘要字数验收。
- 支持旧版 Profile，并提供向后兼容的 v2 Profile。

## 适用场景

### 1. 有官方 DOCX 模板

适合学校或单位提供了正式 `.docx` 模板的情况，例如：

- 封面包含精确下划线、文本框、制表位或浮动对象。
- 前置部分使用罗马页码，正文从阿拉伯数字 1 重新编号。
- 奇偶页使用不同页眉。
- 目录、题注和交叉引用必须保留为 Word 字段。
- 模板中包含声明页、签名区域、内容控件或特殊分节。

此模式以官方模板为权威来源，优先保留模板原有结构，只替换占位内容并把论文内容映射到语义区域，不重建复杂封面坐标。

### 2. 没有模板，只有格式说明

适合只有以下材料的情况：

- 学校发布的文字规范。
- PDF 格式指南。
- 网页说明。
- 格式要求截图。
- 用户直接描述的字号、字体、行距、页码和章节规则。

Codex 会先把要求翻译成 JSON Profile。无法量化、互相矛盾或缺少信息的规则会写入：

```json
{
  "requirements": {
    "unresolved": ["封面填写线长度未给出"],
    "conflicts": []
  }
}
```

只要 `unresolved` 或 `conflicts` 非空，应用脚本默认拒绝执行，防止在不确定条件下破坏文档。

### 3. 已经有可复用 Profile

适合同一学校、学院或期刊的批量文档：

- 首次从模板或规范生成 Profile。
- 后续对多份论文重复执行审计、应用和验收。
- 每份文档只维护少量封面变量、章节定位规则和例外项。

## 仓库结构

```text
.
├── README.md
├── dist/
│   └── standardize-docx-format-v1.0.0.zip
└── standardize-docx-format/
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    ├── scripts/
    │   ├── audit_docx.py
    │   ├── apply_profile.py
    │   └── validate_docx.py
    └── references/
        ├── profile-schema.md
        ├── spec-driven-workflow.md
        ├── ooxml-advanced.md
        ├── thesis-cn.example.json
        └── guangzhou-nanfang-thesis.example.json
```

真正需要安装到 Codex 的目录是 `standardize-docx-format/`。仓库根目录的 README 和 `dist/` 不属于运行时 Skill。

## 环境要求

必需：

- Python 3.10 或更高版本。
- 能够读写本地 `.docx` 文件。
- Codex 或兼容的 Skills 运行环境。

可选：

- Microsoft Word：用于更新全部字段和最终视觉验版。
- LibreOffice：用于无 Word 环境下的兼容渲染检查。
- `docx-cli`：可用于额外的定位编辑、差异比较和渲染；核心脚本不依赖它。

核心 Python 脚本只使用标准库，不要求安装 `python-docx`、`lxml` 或其他第三方包。

## 安装

### 方式一：下载 ZIP

从 Releases 或 `dist/` 下载 `standardize-docx-format-v1.0.0.zip`，解压后把其中的 `standardize-docx-format` 文件夹复制到 Codex Skills 目录。

Windows 默认位置：

```text
C:\Users\你的用户名\.codex\skills\standardize-docx-format
```

macOS/Linux 默认位置：

```text
~/.codex/skills/standardize-docx-format
```

如果设置了 `CODEX_HOME`，则安装到：

```text
$CODEX_HOME/skills/standardize-docx-format
```

### 方式二：克隆仓库

```bash
git clone https://github.com/zwt0204/standardize-docx-format.git
```

Windows PowerShell：

```powershell
$skillsRoot = if ($env:CODEX_HOME) {
  Join-Path $env:CODEX_HOME "skills"
} else {
  Join-Path $env:USERPROFILE ".codex\skills"
}

Copy-Item `
  -LiteralPath ".\standardize-docx-format\standardize-docx-format" `
  -Destination $skillsRoot `
  -Recurse
```

macOS/Linux：

```bash
skills_root="${CODEX_HOME:-$HOME/.codex}/skills"
mkdir -p "$skills_root"
cp -R ./standardize-docx-format/standardize-docx-format "$skills_root/"
```

安装完成后重新启动 Codex 会话，或让运行环境重新扫描 Skills 目录。

## 在 Codex 中使用

可以直接在请求中点名 Skill：

```text
使用 $standardize-docx-format，把 thesis.docx 按 university-template.docx 排版，保留原文件并输出新文件。
```

```text
使用 $standardize-docx-format，根据这份论文格式说明生成 Profile，列出所有未决项，确认后再应用。
```

```text
使用 $standardize-docx-format，审计 thesis.docx 的分节、字体、页码、目录、题注和参考文献格式，不修改文件。
```

```text
使用 $standardize-docx-format，把这一批 DOCX 按 profile.json 标准化，并分别生成验收报告。
```

## 命令行快速开始

下面的命令可以脱离 Codex 单独运行。为避免覆盖原文件，请始终给输出文件使用新路径。

### 1. 审计输入文件

```powershell
python .\standardize-docx-format\scripts\audit_docx.py `
  "D:\docs\thesis.docx" `
  --output ".\before-audit.json"
```

审计内容包括：

- DOCX ZIP 包是否损坏。
- 段落、run、表格和图片数量。
- 样式定义和实际使用情况。
- 直接格式数量。
- 分节、页面尺寸和页边距。
- 页码格式和重启位置。
- 页眉页脚引用关系。
- Word 字段、书签、绘图、文本框和内容控件。
- 中英文字符和单词统计。

### 2. 应用 Profile

```powershell
python .\standardize-docx-format\scripts\apply_profile.py `
  --input "D:\docs\thesis.docx" `
  --output "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json"
```

安全行为：

- 输入路径和输出路径相同时拒绝执行。
- 输出已存在时默认拒绝覆盖。
- Profile 存在未决项或冲突时默认拒绝执行。
- 写入完成后检查 ZIP 包和 XML 可解析性。

只有在明确知道风险时才使用：

```powershell
--force
```

或：

```powershell
--allow-unresolved
```

### 3. 验收输出文件

```powershell
python .\standardize-docx-format\scripts\validate_docx.py `
  --input "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --output ".\validation.json"
```

验收器会输出机器可读 JSON。如果存在格式或内容要求不满足，脚本会使用非零退出码，便于接入批处理和 CI。

## 最小 Profile 示例

```json
{
  "name": "中文本科论文基础格式",
  "requirements": {
    "sourceType": "description",
    "sources": ["学校论文格式规范"],
    "unresolved": [],
    "conflicts": []
  },
  "page": {
    "size": "a4",
    "orientation": "portrait",
    "marginsIn": {
      "top": 0.98,
      "right": 0.79,
      "bottom": 0.98,
      "left": 0.98,
      "header": 0.59,
      "footer": 0.59
    }
  },
  "defaultRun": {
    "latinFont": "Times New Roman",
    "eastAsiaFont": "宋体",
    "complexScriptFont": "Times New Roman",
    "sizePt": 12,
    "updateTheme": true
  },
  "styles": {
    "Normal": {
      "name": "正文",
      "run": {
        "latinFont": "Times New Roman",
        "eastAsiaFont": "宋体",
        "sizePt": 12
      },
      "paragraph": {
        "alignment": "justify",
        "lineSpacing": 1.5,
        "spaceBeforePt": 0,
        "spaceAfterPt": 0,
        "firstLineChars": 2
      }
    }
  },
  "pageNumbers": [
    {"section": 1, "format": "upper-roman", "start": 1},
    {"section": 4, "format": "decimal", "start": 1}
  ],
  "headingNumbering": {
    "enabled": true,
    "styles": ["Heading1", "Heading2", "Heading3"],
    "patterns": ["%1、", "（%2）", "%3."],
    "formats": ["chinese-counting", "chinese-counting", "decimal"],
    "starts": [1, 1, 1],
    "suffix": "space"
  },
  "fields": {
    "updateOnOpen": true
  },
  "validation": {
    "requiredTexts": ["摘要", "Abstract", "目录", "参考文献"],
    "requiredFields": {"TOC": 1, "PAGE": 1},
    "noPlaceholders": ["\\{\\{[^}]+\\}\\}"]
  }
}
```

完整字段说明见：

- `standardize-docx-format/references/profile-schema.md`
- `standardize-docx-format/references/spec-driven-workflow.md`
- `standardize-docx-format/references/ooxml-advanced.md`

## 模板占位符

模板中的占位符可以跨多个 Word run，也可以位于正文、页眉、页脚或文本框中：

```json
{
  "replacements": [
    {
      "find": "{{THESIS_TITLE}}",
      "replace": "多功能控制系统设计",
      "parts": ["document", "headers"],
      "occurrence": "all",
      "required": true
    }
  ]
}
```

对于精确封面，应先在官方模板副本中建立稳定占位符，再替换文本。不要把浮动文本框重建成普通正文段落，否则容易破坏坐标、下划线和对齐。

## 页眉页脚与分节页码

每条规则可以针对一个节和一种页眉页脚类型：

```json
{
  "headersFooters": [
    {
      "section": 4,
      "kind": "footer",
      "type": "default",
      "action": "set",
      "paragraphs": [
        {
          "alignment": "center",
          "segments": [{"field": "PAGE", "result": "1"}],
          "run": {
            "latinFont": "Times New Roman",
            "sizePt": 9
          }
        }
      ]
    }
  ]
}
```

`type` 支持：

- `default`
- `first`
- `even`

创建偶数页页眉或页脚时，脚本会同步启用 Word 的奇偶页设置。

## 语义段落映射

旧论文经常使用手工编号和大量直接格式。可以用有界规则把段落映射到语义样式：

```json
{
  "paragraphRules": [
    {
      "match": {
        "textRegex": "^[一二三四五六七八九十]+、",
        "section": 4,
        "inTextBox": false
      },
      "style": "Heading1",
      "textRegexReplace": {
        "pattern": "^[一二三四五六七八九十]+、\\s*",
        "replacement": ""
      },
      "required": true,
      "maxMatches": 50
    }
  ]
}
```

规则必须尽量限定节、样式、位置或匹配数量，避免使用过宽的全局正则修改整篇论文。

## 验收建议

推荐每次执行完整闭环：

```text
原文件
  ↓
修改前审计
  ↓
生成/确认 Profile
  ↓
输出新 DOCX
  ↓
修改后审计
  ↓
机器验收
  ↓
Word 更新全部字段
  ↓
最终视觉验版
```

重点检查：

- 封面下划线和字段是否仍然对齐。
- 中文摘要、英文摘要、目录和正文是否正确分节。
- 罗马页码是否连续。
- 正文是否从阿拉伯数字 1 重新编号。
- 自动目录是否只包含要求的标题级别。
- 目录页码是否已刷新。
- 奇偶页页眉是否正确。
- 图表、公式和表格是否越界。
- 标题是否出现在页末而正文落到下一页。
- 参考文献、致谢和附录是否按要求另起页。
- 是否残留模板说明或 `{{PLACEHOLDER}}`。

在 Microsoft Word 中可使用：

```text
Ctrl+A → F9
```

更新全部字段。如果 Word 询问目录更新方式，选择“更新整个目录”。

## 已知边界

- Python 脚本可以创建和保留 Word 字段，但不会像 Microsoft Word 一样计算最终字段显示结果。
- 自动目录页码、交叉引用页码和总页数必须由 Word 或兼容渲染器刷新。
- 精确封面依赖官方模板中已有的文本框、形状、坐标和下划线结构。
- 不同 Word 版本、字体安装情况和打印机驱动可能造成轻微分页差异。
- 如果模板要求中文章节号，但图表章号要求阿拉伯数字，需要在 Profile 中明确章号策略，不能直接假设 `STYLEREF` 的显示满足要求。
- CITATION 和 BIBLIOGRAPHY 字段基础设施不等于完整的文献管理器；学校引用体系仍需明确。
- 宏、ActiveX、外部 OLE 对象和第三方插件生成的私有扩展需要单独验证。
- 机器验收不能替代最终视觉验版。

## 数据安全

- 不要把真实论文、身份证明、签名图片、学生信息或未公开模板提交到公共仓库。
- `.gitignore` 默认排除 DOCX、审计报告、验收报告和常见临时文件。
- 核心脚本不会把文档上传到网络，所有处理均在本地完成。
- 输入文件默认只读，输出使用新路径。
- 分享 Profile 前应检查封面字段和替换值中是否包含个人信息。

## 开发与验证

语法编译：

```powershell
python -B -X utf8 -m py_compile `
  .\standardize-docx-format\scripts\audit_docx.py `
  .\standardize-docx-format\scripts\apply_profile.py `
  .\standardize-docx-format\scripts\validate_docx.py
```

JSON 示例解析：

```powershell
Get-ChildItem .\standardize-docx-format\references\*.json | ForEach-Object {
  Get-Content -Raw -Encoding UTF8 $_.FullName | ConvertFrom-Json | Out-Null
}
```

Skill 结构校验可使用 Codex 自带的 `skill-creator/scripts/quick_validate.py`。

## 设计原则

1. 保留原文件，始终输出副本。
2. 先审计，再修改，再验收。
3. 优先修改样式定义，谨慎使用逐段直接格式。
4. 特殊封面和声明页以官方模板为准。
5. 所有自然语言要求先结构化，不确定项必须显式记录。
6. Word 字段保持为字段，不把显示结果硬编码为普通文本。
7. 每条语义规则必须可定位、可审计并尽量有匹配上限。
8. 包结构正确不等于视觉排版完全正确，最终必须渲染验版。

## 反馈与扩展方向

欢迎通过 Issues 提交：

- 新学校或期刊的匿名化 Profile。
- 可复现的 DOCX 结构问题。
- 新的编号格式、字段或题注规则。
- Microsoft Word 与 LibreOffice 的兼容性差异。
- 不包含个人信息的最小复现文件。

建议的后续扩展包括：

- 批量目录和批量验收命令。
- HTML/PDF 格式规范半自动提取器。
- 更完整的按章图表编号策略。
- Word/LibreOffice 自动渲染与页面截图对比。
- Profile JSON Schema 和编辑器自动补全。
- GitHub Actions 自动测试与发布。

## 许可证

当前仓库未附加开源许可证。未经版权所有者明确授权，默认保留全部权利。
