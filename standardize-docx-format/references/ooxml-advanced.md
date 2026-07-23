# 高级 OOXML 标准化约束

## 目录

- [安全边界](#安全边界)
- [脚本字体](#脚本字体)
- [分节页码](#分节页码)
- [标题多级编号](#标题多级编号)
- [页眉页脚关系](#页眉页脚关系)
- [域](#域)
- [封面对齐](#封面对齐)
- [验证](#验证)

## 安全边界

- 优先使用 `docx-cli` 的高层命令。
- 修改前读取精确 XML；修改后进行 ZIP、XML、DOCX schema 和渲染验证。
- 保持元素的 ECMA-376 顺序。错误顺序会导致 Word 报“发现不可读取的内容”。
- 不删除未知关系、部件、扩展命名空间或嵌入对象。
- 不修改 `w:sectPrChange` 中保存的历史快照，除非任务明确要求处理修订。

## 脚本字体

Word 使用不同属性选择不同文字脚本：

```xml
<w:rFonts
  w:ascii="Times New Roman"
  w:hAnsi="Times New Roman"
  w:eastAsia="宋体"
  w:cs="Times New Roman"/>
```

设置显式字体时删除相应的 `asciiTheme`、`hAnsiTheme`、`eastAsiaTheme`、`cstheme`，否则主题字体可能覆盖显式要求。严格论文模板还应同步 `word/theme/theme1.xml` 的 `a:latin`、`a:ea`、`a:cs`。

## 分节页码

在目标节的 `w:sectPr` 中写入：

```xml
<w:pgNumType w:fmt="lowerRoman" w:start="1"/>
```

正文重新从 1 开始：

```xml
<w:pgNumType w:fmt="decimal" w:start="1"/>
```

`w:pgNumType` 必须位于 `w:lnNumType` 之后、`w:cols` 之前。页码显示仍由页眉/页脚中的 PAGE 域完成。

## 标题多级编号

论文编号必须使用组合式 `w:lvlText`：

```xml
<w:lvl w:ilvl="2">
  <w:start w:val="1"/>
  <w:numFmt w:val="decimal"/>
  <w:pStyle w:val="Heading3"/>
  <w:suff w:val="space"/>
  <w:lvlText w:val="%1.%2.%3"/>
  <w:lvlJc w:val="left"/>
</w:lvl>
```

再在 `Heading3` 的 `w:pPr` 中绑定同一个 `numId` 和 `ilvl=2`。不要把编号直接写进标题文本，否则目录、交叉引用和章节重排都会失效。

中文编号由 `w:numFmt` 和 `w:lvlText` 共同决定。例如一级使用 `chineseCounting` + `%1、`，二级使用 `chineseCounting` + `（%2）`，圈码使用 `decimalEnclosedCircle`。括号和顿号不要写入正文文字。

## 页眉页脚关系

- 每个 `w:sectPr` 可分别引用 `default`、`first`、`even` header/footer 部件。
- 缺少某类引用通常表示继承上一节，并不表示空白。要“解除链接并清空”，应创建新的空部件并写入独立关系。
- `first` 需要 `w:titlePg`；`even` 需要 `word/settings.xml` 中的 `w:evenAndOddHeaders`。
- 新建 `word/headerN.xml` 或 `word/footerN.xml` 时，同时更新 `document.xml.rels` 和 `[Content_Types].xml`。
- 学校页眉的横线使用段落下边框，不用空表格或下划线字符。

## 域

常见论文域：

- 目录：`TOC \\o "1-3" \\h \\z \\u`
- 图表题注：`SEQ 图 \\* ARABIC`
- 交叉引用：`REF bookmark \\h`
- 页码引用：`PAGEREF bookmark \\h`
- 引文：`CITATION tag`
- 参考文献表：`BIBLIOGRAPHY`

请求打开时更新域：

```xml
<w:updateFields w:val="true"/>
```

`w:updateFields` 在 `word/settings.xml` 中必须位于 `w:hdrShapeDefaults`、`w:compat`、`w:rsids` 等后续元素之前。CLI 或脚本只能请求更新；最终页码和分页相关结果应由 Microsoft Word 计算。

交叉引用应先创建合法书签，再插入 REF/PAGEREF。Word 书签名使用字母或下划线开头，只包含字母、数字和下划线，最长 40 个字符。不要把 REF 的显示值固定为普通文本。

## 封面对齐

- 优先保留学校模板中的表格、制表位或内容控件。
- 使用固定列宽表格或段落下边框表达填写线。
- 不用重复空格和下划线字符模拟长度；字体替换后它们会漂移。
- 封面通常应作为排除区，避免正文默认字体和段落规则覆盖它。
- 官方模板可能在 `mc:AlternateContent` 中同时保留 DrawingML 与 VML 兼容分支，因此同一个可见文本框可能出现两份 `w:txbxContent`。精确占位符替换应同步命中兼容分支，同时保持锚点、尺寸和未知扩展不变。

## 验证

1. 检查 ZIP 完整性及必需部件。
2. 解析所有被修改的 XML。
3. 使用 `docx validate` 做 ECMA-376 schema 校验。
4. 在 Word 或 LibreOffice 中渲染代表页面。
5. 更新全部域后再次检查目录页码、图表编号和交叉引用。
