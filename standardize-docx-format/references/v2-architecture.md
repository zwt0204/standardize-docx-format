# v2 架构：Document Compiler

v1 已经把正确的底层抽象搭出来了：

```text
自然语言规范 → Profile → Deterministic Engine → DOCX → Validator
```

v2 不在三个脚本上继续堆格式参数，而是补齐 **语义模型、可解释 Diff、可确认 Repair Plan、Visual QA**。LLM 仍然只负责理解规范，OOXML 修改仍然是确定性的。

## 目标流水线

```text
Requirement Source          Official Template
 PDF / 文本 / 网页 / 截图         template.docx
          \                       /
           \                     /
            Requirement IR / Template Profile
                       |
                 Document Spec
                       |
                 JSON Profile
                       |
        ┌──────────────┴──────────────┐
        ↓                             ↓
  Structural Engine            Semantic AST
  (audit / apply)              (document model)
        \                             /
         \                           /
                  Formatting Diff
                          |
                     Repair Plan
                          |
                   用户确认后再 Apply
                          |
                      Renderer
                          |
                     Visual QA
                      /      \
                   PASS     REPAIR
```

## 模块

| 命令 | 脚本 | 职责 |
|---|---|---|
| `audit` | `audit_docx.py` | 只读 OOXML 库存 |
| `model` | `document_model.py` | Document Semantic AST |
| `compile` | `compile_requirements.py` | 规范文本 → Requirement IR + 草稿 Profile |
| `analyze-template` | `analyze_template.py` | 官方模板 → 草稿 Profile |
| `diff` | `diff_profile.py` | expected vs actual，可解释 |
| `plan` | `repair_plan.py` | 把 Diff 变成确认后才能执行的修复步骤 |
| `apply` | `apply_profile.py` | 确定性 OOXML 修改，拒绝原地覆盖 |
| `validate` | `validate_docx.py` | 结构 / 字段 / 字数 / 占位符验收 |
| `visual` | `visual_qa.py` + `layout_estimate.py` | 溢出、孤行、图题分离、分页估算、可选渲染 |
| `repair-visual` | `visual_repair.py` | 保守的 keepNext / 图片缩放，不改正文 |
| `pipeline` | `standardize.py` | 把上面串成一次编译，并写出 `report.html` |

统一入口：

```powershell
python scripts/standardize.py pipeline --input thesis.docx --profile profile.json --work-dir .\qa-output
```

Repair Plan 每一步都有稳定 `id`。确认后可以只执行勾选的步骤：

```powershell
python scripts/standardize.py apply `
  --input thesis.docx `
  --output thesis.standardized.docx `
  --profile profile.json `
  --plan .\qa-output\repair-plan.json `
  --only-auto --yes --force
```

`map_structure` 和 `pageBreakBefore` 不会进入 `--only-auto`；必须把它们的 id 写进 `--only-ids`。

真正写入完整 Profile 仍须显式确认：

```powershell
python scripts/standardize.py pipeline `
  --input thesis.docx `
  --profile profile.json `
  --output thesis.standardized.docx `
  --work-dir .\qa-output `
  --apply --yes --visual-repair
```

## Document Spec

Profile 不再只描述字体字号，也可以描述文档该有哪些语义区域：

```json
{
  "document": {
    "type": "thesis",
    "sections": [
      {"id": "cover", "required": true},
      {"id": "abstract", "required": true, "languages": ["zh", "en"]},
      {"id": "toc", "required": true},
      {"id": "chapter", "required": true, "repeatable": true, "min": 1},
      {"id": "references", "required": true}
    ]
  }
}
```

语义 AST 把段落提升为 `cover / abstract_zh / toc / chapter / figure_caption / references` 等节点。结构规则检查的是“这是不是一篇合格论文”，而不是只检查宋体。

## Diff 与 Repair Plan

Diff 回答：

- 期望什么
- 实际是什么
- 影响多少段落
- 建议怎么修

Repair Plan 把失败项变成可审计步骤。结构缺口（缺摘要、缺参考文献）不会被偷偷改写成正文，而是标记为需要确认的映射问题。

## Visual QA

结构验证不等于视觉验证。Visual QA 现在分三层：

1. **OOXML 启发式**：图片/表格超出正文宽度、标题未 `keepNext`、图题表题分离、连续空段。
2. **分页估算**：用栏宽、字号、行距和图片高度估算页码，发现页末孤行、图题跨页、章标题未另起页。这不是 Word 排版，但能在无渲染器时抓住高风险问题。
3. **可选渲染**：`docx-cli render` 或 LibreOffice → PDF → PNG，检查空白页和转换失败。

保守修复由 `repair-visual` 执行：只设置 `keepNext`、按栏宽缩放图片。`pageBreakBefore` 仍需确认，因为会改变章节分页。

## 不变的安全合同

- 源文件只读，输出新路径。
- Profile 有 `unresolved` / `conflicts` 时默认拒绝 apply。
- 不把字段结果写成普通文本。
- 官方模板的封面、页眉页脚优先 inherit，不重建坐标。
- 视觉闭环可以建议 `keepNext`，但最终分页仍以 Word/LibreOffice 渲染为准。
