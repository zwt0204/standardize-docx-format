# 只有格式说明、没有模板的工作流

## 目录

- [输入归一化](#输入归一化)
- [规则分类](#规则分类)
- [未决项与冲突](#未决项与冲突)
- [语义映射](#语义映射)
- [执行与验收](#执行与验收)

## 输入归一化

接受文字、网页、PDF 或截图中的格式要求。先逐条提取可测量信息，不直接修改 DOCX。

每条规则记录：

- 适用区域：封面、声明、摘要、目录、正文、标题、图表、参考文献、附录、致谢。
- 属性：页面、字体、字号、段落、编号、分页、页眉页脚、域、位置或内容约束。
- 数值与单位：厘米/英寸、磅、倍数、字符缩进、页码起始值。
- 优先级：区域规则高于全局规则；明确例外高于通用规则。
- 来源：规范文件名、页码或条款号。

## 规则分类

将规则写入 profile：

| 规则 | Profile 字段 |
|---|---|
| 全局纸张和页边距 | `page` |
| 特殊章节页面 | `sections` |
| 中文/英文分脚本字体 | `defaultRun`、`styles` |
| 标题和正文识别 | `paragraphRules` |
| 自动标题编号 | `headingNumbering` |
| 罗马/阿拉伯页码 | `pageNumbers` |
| 页眉页脚 | `headersFooters` |
| 封面或模板填写项 | `replacements` |
| 目录、题注、引用 | `fieldRules`、`captionRules`、`bookmarkRules` |
| 字数和章节完整性 | `validation` |

如果规范要求一种尚未建模的视觉构造，保留原文并写入 `requirements.unresolved`，不要用空格或手工编号临时模拟。

## 未决项与冲突

常见未决项：

- “适当”“美观”“按学校标准”等没有测量值。
- 封面填写线只描述文字，没有长度、位置或样例。
- 页码从“正文”开始，但没有定义正文起点或现有 DOCX 没有分节。
- 要求按章图表编号，但标题章号和题注章号使用不同数字体系。
- 同一区域同时出现两个字体、字号或行距要求。

将问题写入：

```json
{
  "requirements": {
    "sourceType": "description",
    "unresolved": ["正文起点尚未定位"],
    "conflicts": ["条款 3 要求一级标题四号，条款 8 要求三号"]
  }
}
```

默认不应用有未决项的 profile。先向用户报告；确认后清空列表并记录采用的解释。

## 语义映射

没有模板时，先审计文档，再为内容建立语义样式：

1. 优先使用已有 Heading/Normal/Caption 等样式。
2. 对无样式段落，根据标题文字模式、所在节、表格/文本框位置建立 `paragraphRules`。
3. 只在规则明确时清除直接格式；正文中的强调、变量、拉丁学名等可能是有效例外。
4. 手写标题编号应先移除，再绑定自动编号。
5. 无法可靠识别的段落输出待复核清单，不做宽泛全局替换。

## 执行与验收

```powershell
python scripts/audit_docx.py input.docx --output before-audit.json
python scripts/apply_profile.py --input input.docx --output output.standardized.docx --profile profile.json
python scripts/audit_docx.py output.standardized.docx --output after-audit.json
python scripts/validate_docx.py --input output.standardized.docx --profile profile.json --output validation.json
```

最后在 Microsoft Word 中：

1. 更新全部域。
2. 检查封面、目录、正文首页、图表页、参考文献和末页。
3. 检查分页、孤行/寡行、浮动对象、表格溢出和交叉引用结果。
4. 保存后再次运行审计，确认字段、书签和分节没有被破坏。
