# 图表审查

- 结果：WARN
- 需要配图章节：4
- 应有图但没有图：4

## 应有图但没有图

- 1.1：work\40_drafts\v1_sections\1.1.md；命中：模型训练
  - 建议：/tender:diagram-gen work\40_drafts\v1_sections\1.1.md --engine fireworks --insert
- 1.2：work\40_drafts\v1_sections\1.2.md；命中：模型训练
  - 建议：/tender:diagram-gen work\40_drafts\v1_sections\1.2.md --engine fireworks --insert
- 1.3：work\40_drafts\v1_sections\1.3.md；命中：模型训练
  - 建议：/tender:diagram-gen work\40_drafts\v1_sections\1.3.md --engine fireworks --insert
- 1.4：work\40_drafts\v1_sections\1.4.md；命中：模型训练
  - 建议：/tender:diagram-gen work\40_drafts\v1_sections\1.4.md --engine fireworks --insert

处理要求：review 必须在生成 review.md 之前补图；如果无法调用 /tender:diagram-gen，必须执行同等逻辑生成 SVG 并插入章节 Markdown。
