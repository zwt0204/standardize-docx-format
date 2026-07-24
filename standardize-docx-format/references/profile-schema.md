# DOCX 标准化 Profile v2

## 目录

- [顶层结构](#顶层结构)
- [规范来源与安全门](#规范来源与安全门)
- [页面与分节](#页面与分节)
- [字体和样式](#字体和样式)
- [多级编号](#多级编号)
- [语义段落规则](#语义段落规则)
- [模板占位符](#模板占位符)
- [页眉页脚](#页眉页脚)
- [域书签和题注](#域书签和题注)
- [验收规则](#验收规则)

## 顶层结构

使用 UTF-8 JSON。所有 v1 字段继续兼容。

```json
{
  "name": "标准名称",
  "requirements": {},
  "page": {},
  "sections": [],
  "defaultRun": {},
  "styles": {},
  "headingNumbering": {},
  "pageNumbers": [],
  "paragraphRules": [],
  "captionRules": [],
  "bookmarkRules": [],
  "headersFooters": [],
  "replacements": [],
  "fieldRules": [],
  "fields": {},
  "validation": {}
}
```

执行顺序：全局页面 → 逐节覆盖 → 样式/编号 → 语义段落 → 题注/书签 → 页眉页脚 → 文本替换 → 域 → 域刷新。

## 规范来源与安全门

没有模板、只有文字规范时，由 Agent 将要求翻译为 profile，并记录无法量化或冲突的条款：

```json
{
  "requirements": {
    "sourceType": "description",
    "sources": ["学校论文写作规范 2026"],
    "unresolved": ["封面填写线长度未给出"],
    "conflicts": []
  }
}
```

`unresolved` 或 `conflicts` 非空时，`apply_profile.py` 默认拒绝执行。只有经过人工确认后清空，或显式传入 `--allow-unresolved` 才继续。

## 页面与分节

`page` 应用到所有节，`sections[].page` 随后覆盖指定节。

```json
{
  "page": {
    "size": "a4",
    "orientation": "portrait",
    "marginsIn": {
      "top": 1.0,
      "right": 1.0,
      "bottom": 1.0,
      "left": 1.25,
      "header": 0.6,
      "footer": 0.6,
      "gutter": 0
    }
  },
  "sections": [
    {
      "section": 3,
      "type": "nextPage",
      "differentFirstPage": false,
      "page": {"marginsIn": {"top": 1.2}}
    }
  ]
}
```

页面尺寸：`a3`、`a4`、`a5`、`letter`、`legal`、`tabloid`，或同时指定 `widthIn`/`heightIn`。

节类型：`nextPage`、`continuous`、`evenPage`、`oddPage`、`nextColumn`。

分节页码：

```json
{
  "pageNumbers": [
    {"section": 1, "format": "upper-roman", "start": 1},
    {"section": 4, "format": "decimal", "start": 1}
  ]
}
```

页码格式：`decimal`、`lower-roman`、`upper-roman`、`lower-alpha`、`upper-alpha`。

## 字体和样式

`defaultRun` 和 `styles.*.run` 支持：

- `latinFont` → `w:ascii`、`w:hAnsi`
- `eastAsiaFont` → `w:eastAsia`
- `complexScriptFont` → `w:cs`
- `sizePt`、`bold`、`italic`、`color`
- `updateTheme` 仅用于 `defaultRun`

```json
{
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
      "run": {"latinFont": "Times New Roman", "eastAsiaFont": "宋体", "sizePt": 12},
      "paragraph": {
        "alignment": "justify",
        "lineSpacing": 1.5,
        "spaceBeforePt": 0,
        "spaceAfterPt": 0,
        "firstLineChars": 2,
        "widowControl": true
      }
    }
  }
}
```

段落属性：`alignment`、`lineSpacing`、`lineSpacingExactPt`、`spaceBeforePt`、`spaceAfterPt`、`leftIn`、`rightIn`、`firstLineIn`、`hangingIn`、`firstLineChars`、`keepNext`、`keepLines`、`pageBreakBefore`、`widowControl`、`outlineLevel`。

## 多级编号

```json
{
  "headingNumbering": {
    "enabled": true,
    "styles": ["Heading1", "Heading2", "Heading3", "Heading4", "Heading5"],
    "patterns": ["%1、", "（%2）", "%3.", "（%4）", "%5"],
    "formats": [
      "chinese-counting",
      "chinese-counting",
      "decimal",
      "decimal",
      "decimal-enclosed-circle"
    ],
    "starts": [1, 1, 1, 1, 1],
    "suffix": "space"
  }
}
```

额外格式：`chinese-counting-thousand`、`chinese-legal-simplified`、`ideograph-traditional`、`ideograph-digital`、`decimal-enclosed-fullstop`、`decimal-full-width`。

`patterns` 负责括号、顿号、点号等显示字符；编号必须绑定样式，不要写死在标题文字中。

## 语义段落规则

对直接格式严重、没有正确样式的文档，使用 `paragraphRules` 映射语义：

```json
{
  "paragraphRules": [
    {
      "match": {"textRegex": "^第?[一二三四五六七八九十]+[、.]", "section": 4, "inTextBox": false},
      "style": "Heading1",
      "clearParagraphFormatting": true,
      "textRegexReplace": {"pattern": "^第?[一二三四五六七八九十]+[、.]\\s*", "replacement": ""},
      "required": true
    },
    {
      "match": {"textContains": "填写说明"},
      "remove": true
    }
  ]
}
```

匹配条件：`textEquals`、`textContains`、`startsWith`、`textRegex`、`textNotRegex`、`currentStyle`、`currentStyleNotIn`、`styleNameRegex`、`styleNameNotRegex`、`section`、`paragraphIndex`、`inTable`、`inTextBox`、`normalizeWhitespace`。

TOC / 目录保护建议：

- 真实中文论文几乎都有 `toc 1` / `toc 2` 目录行，且文本形态与章节标题相似（如 `1 绪论 1`、`1.1 研究背景 1`）。
- 对“数字开头即标题”的规则，务必排除目录样式：`styleNameNotRegex: "(?i)^toc"`。
- 目录条目常以页码结尾；可用 `textNotRegex: ".+\d$"` 降低误伤（在 `normalizeWhitespace` 后生效）。
- 若正文标题本身已带手工编号，而 profile 又启用了 `headingNumbering`，可用 `textRegexReplace` 去掉正文标题前缀编号；**不要**对目录行做同样替换。


动作：`style`、`paragraph`、`run`、`clearParagraphFormatting`、`clearRunFormatting`、`textRegexReplace`、`remove`、`required`、`maxMatches`。

## 模板占位符

`replacements` 可替换跨 run 的文本，并保留文本框、VML/DrawingML 锚点和原格式：

```json
{
  "replacements": [
    {
      "find": "{{THESIS_TITLE}}",
      "replace": "数字化转型背景下的……研究",
      "parts": ["document", "headers"],
      "occurrence": "all",
      "required": true
    }
  ]
}
```

`parts`：`document`、`headers`、`footers`、`notes`、`all` 或精确 XML 部件路径。支持 `regex: true`；正则替换使用 Python replacement 语法。

## 页眉页脚

每条规则处理一个节、一个类型：

```json
{
  "headersFooters": [
    {
      "section": 4,
      "kind": "header",
      "type": "default",
      "action": "set",
      "paragraphs": [
        {
          "alignment": "center",
          "segments": [{"text": "广州南方学院成人高等学历继续教育本科毕业论文（设计）"}],
          "run": {"eastAsiaFont": "宋体", "latinFont": "Times New Roman", "sizePt": 9},
          "bottomBorder": {"style": "single", "size": 6, "space": 1, "color": "000000"}
        }
      ]
    },
    {
      "section": 4,
      "kind": "footer",
      "type": "default",
      "action": "set",
      "paragraphs": [
        {
          "alignment": "center",
          "segments": [{"field": "PAGE", "result": "1"}],
          "run": {"latinFont": "Times New Roman", "sizePt": 9}
        }
      ]
    }
  ]
}
```

`type`：`default`、`first`、`even`。`action`：`set`、`clear`、`inherit`。创建 `even` 时会启用 Word 的奇偶页页眉设置。

段落 `segments` 支持普通 `text` 或安全 Word `field`。字段白名单包括 PAGE、NUMPAGES、SECTIONPAGES、TOC、SEQ、REF、PAGEREF、STYLEREF、CITATION、BIBLIOGRAPHY、NOTEREF、HYPERLINK。

## 域书签和题注

用文本占位符创建域：

```json
{
  "fieldRules": [
    {"placeholder": "{{TOC}}", "instruction": "TOC \\o \"1-2\" \\h \\u", "result": "目录", "required": true},
    {"placeholder": "{{REF_FIG_1}}", "instruction": "REF fig_1 \\h", "result": "图1-1"},
    {"placeholder": "{{REF_PAGE_FIG_1}}", "instruction": "PAGEREF fig_1 \\h", "result": "1"},
    {"placeholder": "{{BIBLIOGRAPHY}}", "instruction": "BIBLIOGRAPHY", "result": "参考文献"}
  ],
  "fields": {"updateOnOpen": true}
}
```

书签规则：

```json
{
  "bookmarkRules": [
    {"name": "chapter_{index}", "match": {"currentStyle": "Heading1"}, "maxMatches": 9}
  ]
}
```

题注规则会删除原手工编号，创建 SEQ 域，可选章节域和书签：

```json
{
  "captionRules": [
    {
      "label": "图",
      "sequence": "图",
      "match": {"textRegex": "^图\\s*\\d+(?:[-.]\\d+)?"},
      "stripPattern": "^图\\s*\\d+(?:[-.]\\d+)?\\s*",
      "style": "FigureCaption",
      "chapterFieldInstruction": "STYLEREF \"Heading 1\" \\n",
      "separator": "-",
      "bookmarkPrefix": "fig_"
    }
  ]
}
```

如果学校的一级标题使用中文序号但图表要求阿拉伯章号，应在 `requirements.unresolved` 中记录并选择专门的章号字段策略，不能假定 STYLEREF 的显示一定符合要求。

## 验收规则

运行：

```powershell
python scripts/validate_docx.py --input output.docx --profile profile.json --output validation.json
```

```json
{
  "validation": {
    "sectionCount": 6,
    "requiredTexts": [
      {"pattern": "^摘\\s*要$", "regex": true, "min": 1},
      "参考文献"
    ],
    "forbiddenTexts": ["填写说明", "此处删除"],
    "noPlaceholders": ["\\{\\{[^}]+\\}\\}"],
    "requiredFields": {"TOC": 1, "PAGE": 1},
    "requiredBookmarks": ["fig_1"],
    "requiredStylesUsed": ["Heading1", "Normal"],
    "wordCounts": [
      {
        "name": "中文摘要字数",
        "startRegex": "^摘\\s*要$",
        "endRegex": "^关键词[:：]",
        "mode": "cjkCharacters",
        "min": 300,
        "max": 500
      },
      {
        "name": "英文摘要词数",
        "startRegex": "^Abstract$",
        "endRegex": "^Key\\s*words?[:：]",
        "mode": "latinWords",
        "min": 250,
        "max": 400
      }
    ]
  }
}
```

校验会同时核对 profile 中的分节页码和新建页眉页脚引用。脚本校验不能替代 Word 的最终分页、浮动对象和域结果渲染检查。
