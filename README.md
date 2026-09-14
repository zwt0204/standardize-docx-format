# standardize-docx-format

[简体中文](README.md) | [English](README.en.md)

面向 Codex 的 Microsoft Word `.docx` 格式标准化 Skill，重点解决毕业论文、学位论文、课程论文、技术报告等文档的复杂排版、模板套用和自动验收问题。

它既可以根据学校提供的官方 Word 模板工作，也可以把自然语言、PDF、网页或截图中的格式要求转换为结构化 JSON Profile，再确定性地应用到现有 DOCX 文件中。

当前实现是 **v2.1**：同一套确定性 OOXML 引擎上，补了 Profile JSON Schema、可勾选的 Repair Plan、更深的模板反推，以及结构验收 / Visual QA。

```text
规范 / 模板 → Requirement IR / Document Spec → Profile
      → Audit / Semantic AST → Diff → Repair Plan
      → 确认后 Apply → Validate → Visual QA
```

LLM 负责理解规范；OOXML 修改仍然是确定性的。默认先给出 Diff 和 Repair Plan，不静默改文档。

> 这个项目首先是一个 Codex Skill，不是 Word 插件，也不是带图形界面的排版软件。核心脚本只依赖 Python 标准库，可以独立完成 DOCX 包审计、语义建模、Diff、Repair Plan、Profile 应用、结构验收和视觉启发式检查；复杂需求识别、模板区域映射和未决规则处理由 Codex 工作流负责。最终分页仍以 Word 或 LibreOffice 渲染为准。

Skill 运行时说明（英文，给 Codex 读）见 [`standardize-docx-format/SKILL.md`](standardize-docx-format/SKILL.md)。

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
- 支持首页、普通页、偶数页分别设置页眉页脚；官方模板可 `action: inherit`，不重建封面坐标、不删除已有页眉页脚引用。
- 支持中文论文多级自动编号：`一、`、`（一）`、`1.`、`（1）`、`①`。
- 支持自动目录、PAGE、NUMPAGES、SEQ、REF、PAGEREF、STYLEREF、CITATION、BIBLIOGRAPHY 等 Word 字段。
- 支持书签、题注、图表编号和交叉引用基础设施；题注章号策略见 `captionNumbering`。
- 支持跨多个 run、页眉、页脚和文本框的精确占位符替换。
- 支持按文本、正则、现有样式、节、段落位置、表格或文本框位置映射语义样式；目录样式可用 `styleNameNotRegex` 保护。
- 支持删除模板教学说明、清理手写标题编号和统一正文格式。
- 支持结构、必备章节、字段、书签、页码、页眉页脚、占位符和中英文摘要字数验收。
- 支持把论文提升为 Document Semantic AST：封面、摘要、目录、章、图题、表题、参考文献。
- 支持 Profile Diff 和 Repair Plan：先解释 expected vs actual，用户确认后再 apply；可只执行 `--only-auto` 或指定 `--only-ids`。
- 支持 Visual QA 三层：OOXML 启发式、分页估算、可选 `docx-cli` / LibreOffice 渲染。
- 支持从规范文本/PDF 草稿编译 Profile，以及从官方模板反推 numbering / TOC / 占位符 / 有界 `paragraphRules`。
- 支持用 `references/profile.schema.json` 校验 Profile。
- 支持旧版 Profile，并提供向后兼容的 v2 Profile / Document Spec。

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
├── README.en.md
├── .github/workflows/ci.yml
├── dist/
│   └── standardize-docx-format-v1.0.0.zip   # 历史打包文件名；以当前源码为准
└── standardize-docx-format/                 # 真正安装到 Codex 的 Skill 目录
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    ├── scripts/
    │   ├── standardize.py          # 统一 CLI / pipeline
    │   ├── profile_schema.py       # Profile JSON Schema 校验
    │   ├── plan_apply.py           # Repair Plan → 过滤后的 Profile
    │   ├── audit_docx.py
    │   ├── document_model.py
    │   ├── compile_requirements.py
    │   ├── analyze_template.py
    │   ├── diff_profile.py
    │   ├── repair_plan.py
    │   ├── apply_profile.py
    │   ├── validate_docx.py
    │   ├── visual_qa.py
    │   ├── layout_estimate.py      # 无渲染器时的分页估算
    │   ├── visual_repair.py        # 保守 keepNext / 图片缩放
    │   ├── report.py               # 合规 HTML 报告
    │   ├── test_paragraph_match_safety.py
    │   └── test_v2_pipeline.py
    └── references/
        ├── v2-architecture.md
        ├── profile.schema.json
        ├── profile-schema.md
        ├── spec-driven-workflow.md
        ├── ooxml-advanced.md
        ├── thesis-cn.example.json
        └── guangzhou-nanfang-thesis.example.json
```

根目录的 README 和 `dist/` 不属于运行时 Skill。GitHub Actions 会编译脚本、校验示例 Profile 并跑回归测试。

## 环境要求

必需：

- Python **3.10 或更高**。CI 目前在 3.11 / 3.12 上跑；请不要用 3.9 及以下。
- 能够读写本地 `.docx` 文件。
- 作为 Codex Skill 使用时，需要 Codex 或兼容的 Skills 运行环境。命令行可单独跑脚本。

可选：

- Microsoft Word：更新全部字段和最终视觉验版（金标准）。
- LibreOffice / `soffice`：无 Word 时把 DOCX 转 PDF。
- `pdftoppm`（poppler）：把 PDF 切成 PNG，供 Visual QA 和人工/agent 看页。
- `docx-cli`（`docx` 命令）：定位编辑、差异比较和渲染；核心脚本不依赖它。有则优先用 `docx render`。

核心 Python 脚本只使用标准库，不要求安装 `python-docx`、`lxml` 或其他第三方包。

## 安装

### 方式一：下载 ZIP

从 Releases 或 `dist/` 下载 zip。当前仓库里的打包文件仍可能叫 `standardize-docx-format-v1.0.0.zip`，**内容请以本仓库 `standardize-docx-format/` 源码为准**，不要只凭文件名判断版本。解压后把其中的 `standardize-docx-format` 文件夹复制到 Codex Skills 目录。

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

下面的命令可以脱离 Codex 单独运行。为避免覆盖原文件，请始终给输出文件使用新路径。Windows 示例用 PowerShell；Linux/macOS 把反引号续行换成 `\`，路径按本机修改。

### 1. 审计输入文件

```powershell
python .\standardize-docx-format\scripts\audit_docx.py `
  "D:\docs\thesis.docx" `
  --output ".\before-audit.json"
```

或：

```powershell
python .\standardize-docx-format\scripts\standardize.py audit `
  --input "D:\docs\thesis.docx" `
  --output ".\before-audit.json"
```

审计内容包括：DOCX ZIP 是否损坏；段落 / run / 表格 / 图片数量；样式定义与使用；直接格式；分节与页边距；页码格式和重启；页眉页脚引用；字段、书签、绘图、文本框、内容控件；中英文字符和单词统计。

### 2. 校验 Profile，再看 Diff / Repair Plan

```powershell
python .\standardize-docx-format\scripts\standardize.py schema --input ".\profile.json"
python .\standardize-docx-format\scripts\standardize.py diff `
  --input "D:\docs\thesis.docx" --profile ".\profile.json" --text
python .\standardize-docx-format\scripts\standardize.py plan `
  --input "D:\docs\thesis.docx" --profile ".\profile.json" --text
```

未确认前不要 `apply`。Repair Plan 里：

- `--only-auto` 只跑可自动步骤。
- `map_structure`（结构缺口，例如缺摘要）和 `pageBreakBefore` **不会**进入 auto，必须 `--only-ids`。
- 有 `--plan` 时必须同时给 `--yes`，否则拒绝写入。

### 3. 应用 Profile

完整 Profile：

```powershell
python .\standardize-docx-format\scripts\apply_profile.py `
  --input "D:\docs\thesis.docx" `
  --output "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json"
```

只执行已确认的自动修复步骤：

```powershell
python .\standardize-docx-format\scripts\standardize.py apply `
  --input "D:\docs\thesis.docx" `
  --output "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --plan ".\qa-output\repair-plan.json" `
  --only-auto --yes --force
```

安全行为：

- 输入路径和输出路径相同时拒绝执行。
- 输出已存在时默认拒绝覆盖（需要时才加 `--force`）。
- Profile 存在未决项或冲突时默认拒绝执行（需要时才加 `--allow-unresolved`）。
- 写入完成后检查 ZIP 包和 XML 可解析性。

### 4. 结构验收

```powershell
python .\standardize-docx-format\scripts\validate_docx.py `
  --input "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --output ".\validation.json"
```

不满足要求时非零退出，便于 CI。结构验收通过 **不等于** 视觉排版通过。

### 5. Visual QA 与真实渲染

默认只跑 OOXML 启发式和分页估算，不调用 Word：

```powershell
python .\standardize-docx-format\scripts\standardize.py visual `
  --input "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --output ".\qa-output\visual.json" `
  --html ".\qa-output\visual-report.html"
```

要出 PDF/PNG，必须给 `--render-dir`（`standardize.py visual` **没有** `--render` 开关）：

```powershell
python .\standardize-docx-format\scripts\standardize.py visual `
  --input "D:\docs\thesis.standardized.docx" `
  --profile ".\profile.json" `
  --render-dir ".\qa-output\rendered" `
  --output ".\qa-output\visual.json" `
  --html ".\qa-output\visual-report.html"
```

渲染优先级：

1. 若存在 `docx` 命令：`docx render … --out <render-dir>/rendered`
2. 否则 LibreOffice `--headless --convert-to pdf`；若有 `pdftoppm` 再切 `page*.png`
3. 都没有则 `render.skipped=true`，启发式报告仍会写出

机器目前只会标「PNG 相对中位过小（可能空白页）」和「PDF 过小（可能转换失败）」。页眉叠字、表格溢出、封面位移、目录页码、中文编号、STYLEREF **必须看图或在 Word 里看**。

有 `docx-cli` 时也可直接：

```powershell
docx render output.standardized.docx --out rendered-pages
```

人工至少看五页：封面、目录、正文第一页、有图或表的一页、最后一页。然后在 Word 中 `Ctrl+A` → `F9` 刷新域（目录选“更新整个目录”）。LibreOffice 过了不能保证 Word 里 TOC 页码和 STYLEREF 一致。

### 6. v2 编译器流水线

先看差异和修复计划，确认后再写入：

```powershell
python .\standardize-docx-format\scripts\standardize.py compile `
  --input ".\spec.txt" `
  --output ".\requirement-ir.json" `
  --profile-out ".\profile.json"

python .\standardize-docx-format\scripts\standardize.py analyze-template `
  --input ".\university-template.docx" `
  --profile-out ".\template-profile.json"

python .\standardize-docx-format\scripts\standardize.py pipeline `
  --input "D:\docs\thesis.docx" `
  --profile ".\profile.json" `
  --work-dir ".\qa-output"

python .\standardize-docx-format\scripts\standardize.py pipeline `
  --input "D:\docs\thesis.docx" `
  --profile ".\profile.json" `
  --output "D:\docs\thesis.standardized.docx" `
  --work-dir ".\qa-output" `
  --apply --yes --plan-apply --only-auto --render
```

`pipeline --render` 会尝试写出 `qa-output/rendered-before/`（apply 前）和 `qa-output/rendered/`（apply 后）。`--visual-repair` 只做保守的 `keepNext` / 按栏宽缩放图片。

工作目录常见产物：

```text
qa-output/
  before-audit.json
  after-audit.json          # 仅 --apply
  document-model.json
  diff.json
  repair-plan.json
  apply.json                # 仅 --apply
  validation.json           # 仅 --apply
  visual-before.json
  visual.json
  visual-report.html
  report.html
  compliance.json
  pipeline-report.json
  rendered-before/          # 仅 --render
  rendered/                 # 仅 --render 且 --apply
```

其它子命令：

```powershell
python .\standardize-docx-format\scripts\standardize.py model --input thesis.docx --output model.json
python .\standardize-docx-format\scripts\standardize.py repair-visual --input thesis.docx --output thesis.visual.docx --profile profile.json
```

## 论文目录（TOC）安全

真实中文论文几乎都有 `toc 1` / `toc 2` 目录行，文本形态常和章节标题相似（如 `1 绪论 1`）。对“数字开头即标题”的规则必须先排除目录：

```json
{
  "match": {
    "textRegex": "^\\d+\\s+\\S+",
    "styleNameNotRegex": "(?i)^toc",
    "textNotRegex": ".+\\d$",
    "inTextBox": false
  },
  "style": "Heading1",
  "textRegexReplace": {"pattern": "^\\d+\\s+", "replacement": ""},
  "maxMatches": 50
}
```

匹配条件还包括 `textNotRegex`、`currentStyleNotIn`、`styleNameRegex`、`styleNameNotRegex`；样式显示名从 `word/styles.xml` 解析。`references/thesis-cn.example.json` 和 `guangzhou-nanfang-thesis.example.json` 已带 TOC 保护。前者只是起点，不是某所学校的最终规范。

## 题注章号 `captionNumbering`

中文章节号（`一、`）和图表阿拉伯章号（`图 1-1`）经常不一致，不能假定 `STYLEREF` 显示值就是学校要的。

```json
{
  "captionNumbering": {
    "chapterStyle": "Heading 1",
    "headingDisplay": "chinese-counting",
    "chapterDisplay": "decimal",
    "separator": "-",
    "strategy": "styleref-arabic"
  }
}
```

| `strategy` | 含义 |
|---|---|
| `styleref-as-displayed` | 题注章号与标题编号显示一致 |
| `styleref-arabic` | 标题可为中文序号，题注章号强制阿拉伯数字 |
| `seq-chapter` | 独立 `SEQ chapter` |
| `needs-confirmation` | 必须留在 `requirements.unresolved`，默认拒绝 apply |

`captionRules` 未写 `chapterFieldInstruction` 时使用该策略。

## 最小 Profile 示例

```json
{
  "name": "中文本科论文基础格式",
  "profileVersion": "2.1",
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
  "captionNumbering": {
    "chapterStyle": "Heading 1",
    "separator": "-",
    "strategy": "styleref-as-displayed"
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

- [`standardize-docx-format/references/profile.schema.json`](standardize-docx-format/references/profile.schema.json)
- [`standardize-docx-format/references/profile-schema.md`](standardize-docx-format/references/profile-schema.md)
- [`standardize-docx-format/references/spec-driven-workflow.md`](standardize-docx-format/references/spec-driven-workflow.md)
- [`standardize-docx-format/references/ooxml-advanced.md`](standardize-docx-format/references/ooxml-advanced.md)
- [`standardize-docx-format/references/v2-architecture.md`](standardize-docx-format/references/v2-architecture.md)

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

`type`：`default`、`first`、`even`。`action`：`set`、`clear`、`inherit`。

- `set`：写入指定段落。
- `clear`：清空该槽。
- `inherit`：保留模板已有页眉页脚，**不删除引用**。官方封面/页眉优先用这个。

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
        "inTextBox": false,
        "styleNameNotRegex": "(?i)^toc"
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
原文件 / 规范 / 模板
  ↓
Requirement IR + Document Spec
  ↓
修改前审计 + Semantic AST
  ↓
Diff + Repair Plan
  ↓
用户确认
  ↓
输出新 DOCX
  ↓
结构验收 + Visual QA（可选渲染 PNG/PDF）
  ↓
Word 更新全部字段
  ↓
最终渲染验版
```

重点检查：

- 封面下划线和字段是否仍然对齐。
- 中文摘要、英文摘要、目录和正文是否正确分节。
- 罗马页码是否连续；正文是否从阿拉伯数字 1 重新编号。
- 目录行仍是 TOC 样式，没有被映射成 Heading。
- 自动目录是否只包含要求的标题级别；页码是否已在 Word 中刷新。
- 奇偶页页眉是否正确。
- 图表、公式和表格是否越界；图题是否跟图表拆页。
- 标题是否出现在页末而正文落到下一页。
- 参考文献、致谢和附录是否按要求另起页。
- 是否残留模板说明或 `{{PLACEHOLDER}}`。

`validate` ok、`visual.ok`、ZIP 能打开只说明结构层过了。`render.skipped` 表示这台机器没渲出来，不是视觉通过。

## 已知边界

- Python 脚本可以创建和保留 Word 字段，但不会像 Microsoft Word 一样计算最终字段显示结果。
- 自动目录页码、交叉引用页码和总页数必须由 Word 或兼容渲染器刷新。
- 精确封面依赖官方模板中已有的文本框、形状、坐标和下划线结构。
- 不同 Word 版本、字体安装情况和打印机驱动可能造成轻微分页差异。
- LibreOffice / `docx-cli` 渲染是近似结果，尤其是 CJK 编号、STYLEREF 和 TOC 页码。
- 如果模板要求中文章节号，但图表章号要求阿拉伯数字，需要在 Profile 中明确 `captionNumbering.strategy`。
- CITATION 和 BIBLIOGRAPHY 字段基础设施不等于完整的文献管理器。
- 宏、ActiveX、外部 OLE 对象和第三方插件生成的私有扩展需要单独验证。
- 仓库回归测试目前是合成 OOXML，不是学校脱敏金标论文；CI 也不跑 Word。
- 机器验收不能替代最终视觉验版。

## 数据安全

- 不要把真实论文、身份证明、签名图片、学生信息或未公开模板提交到公共仓库。
- `.gitignore` 默认排除 DOCX、审计报告、验收报告和常见临时文件。
- 核心脚本不会把文档上传到网络，所有处理均在本地完成。
- 输入文件默认只读，输出使用新路径。
- 分享 Profile 前应检查封面字段和替换值中是否包含个人信息。

## 开发与验证

语法编译（与 CI 列表对齐）：

```powershell
python -B -X utf8 -m py_compile `
  .\standardize-docx-format\scripts\audit_docx.py `
  .\standardize-docx-format\scripts\apply_profile.py `
  .\standardize-docx-format\scripts\validate_docx.py `
  .\standardize-docx-format\scripts\document_model.py `
  .\standardize-docx-format\scripts\diff_profile.py `
  .\standardize-docx-format\scripts\repair_plan.py `
  .\standardize-docx-format\scripts\visual_qa.py `
  .\standardize-docx-format\scripts\layout_estimate.py `
  .\standardize-docx-format\scripts\visual_repair.py `
  .\standardize-docx-format\scripts\compile_requirements.py `
  .\standardize-docx-format\scripts\analyze_template.py `
  .\standardize-docx-format\scripts\standardize.py `
  .\standardize-docx-format\scripts\profile_schema.py `
  .\standardize-docx-format\scripts\plan_apply.py `
  .\standardize-docx-format\scripts\report.py
```

回归测试：

```powershell
python -B -X utf8 .\standardize-docx-format\scripts\test_paragraph_match_safety.py
python -B -X utf8 .\standardize-docx-format\scripts\test_v2_pipeline.py
```

校验示例 Profile：

```powershell
python -B -X utf8 .\standardize-docx-format\scripts\profile_schema.py `
  --input .\standardize-docx-format\references\thesis-cn.example.json
python -B -X utf8 .\standardize-docx-format\scripts\profile_schema.py `
  --input .\standardize-docx-format\references\guangzhou-nanfang-thesis.example.json
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
4. 特殊封面和声明页以官方模板为准；页眉页脚默认 inherit。
5. 所有自然语言要求先结构化，不确定项必须显式记录。
6. Word 字段保持为字段，不把显示结果硬编码为普通文本。
7. 每条语义规则必须可定位、可审计并尽量有匹配上限；目录样式不得被标题规则误伤。
8. 包结构正确不等于视觉排版完全正确，最终必须渲染验版。
9. Repair Plan 不得静默 apply；`map_structure` 和 `pageBreakBefore` 必须显式勾选。

## 反馈与扩展方向

欢迎通过 Issues 提交：

- 新学校或期刊的匿名化 Profile。
- 可复现的 DOCX 结构问题。
- 新的编号格式、字段或题注规则。
- Microsoft Word 与 LibreOffice 的兼容性差异。
- 不包含个人信息的最小复现文件。

建议的后续扩展包括：

- 用 LibreOffice/Word 渲染页做空白页、页眉碰撞和跨页表检测（启发式仍标为 estimated）。
- 截图 OCR 的 Requirement Compiler（HTML 文本层已支持）。
- 学校 Profile 库与更多匿名金标论文。
- 公式、脚注、表格样式模型。
- 浏览器端预览（核心 Engine 稳定后再做）。
- 将 `dist/` 打包名与当前 v2.1 源码对齐。

## 许可证

当前仓库未附加开源许可证。未经版权所有者明确授权，默认保留全部权利。
