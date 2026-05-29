# 下一步行动

## 当前状态
- 结论：MANUAL
- 当前不能执行 /tender:build：章节完成度不足（4/13），页数估算低于目标下限，page_plan.yml allocation 合计低于85页下限
- 下一条建议命令：/tender:draft

## 需要用户补充
- 智能体服务来源（自研/外购/已有）：接口设计章节需要确认智能体来源，标注为“可后填/不阻断技术正文”
- 图关系模型实现方式（自研/开源）：服务转化模型需要确认，标注为“可后填/不阻断技术正文”
- 与主标接口约定（数据交换协议）：副标接入主标预测结果的接口协议，标注为“可后填/不阻断技术正文”

## P0 先修阻断
- [ ] 扩充 page_plan.yml allocation 合计达到85页下限；目标文件：work/30_plan/page_plan.yml；完成判定：allocation 合计 >= 85页
- [ ] 继续写作第2章（总体技术方案）2.1~2.5；目标文件：work/40_drafts/v1_sections/2.1~2.5.md；完成判定：5个章节均已draft，字数达到目标

## P1 继续写作/扩写
- [ ] 继续写作第3章（详细技术方案）3.1.1~3.1.3；目标文件：work/40_drafts/v1_sections/3.1.x.md；完成判定：3个章节均已draft
- [ ] 补图：1.1项目背景架构图；目标文件：work/40_drafts/v1_sections/1.1.md；完成判定：插入 fireworks SVG 并记录路径到 work/50_figures/
- [ ] 补图：1.2能力体系图；目标文件：work/40_drafts/v1_sections/1.2.md；完成判定：插入 fireworks SVG 并记录路径
- [ ] 补图：1.3建设内容结构图；目标文件：work/40_drafts/v1_sections/1.3.md；完成判定：插入 fireworks SVG 并记录路径
- [ ] 补图：1.4技术重点难点图；目标文件：work/40_drafts/v1_sections/1.4.md；完成判定：插入 fireworks SVG 并记录路径

## P2 表达和格式优化
- [ ] 删除 B4.lock 孤立锁；目标文件：work/40_drafts/locks/B4.lock；完成判定：文件不存在
- [ ] /tender:outline 确认3.3节评分项（3.3.1平台功能架构、3.3.3场景服务分析、3.3.4 GIS空间分析）是否在技术方案范围内；完成判定：scoring_title_placement.md 无 FAIL 项
- [ ] 继续写作第4~7章（实施/培训/售后/综合）；目标文件：work/40_drafts/v1_sections/4.x~7.x.md；完成判定：章节均已draft

## 复查
- [ ] 重新执行 /tender:review
- [ ] review 通过后再执行 /tender:build
