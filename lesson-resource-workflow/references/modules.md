# 模块目录
版本：0.5.4，2026-11-01。

| 顺序 | 模块 | 规范与工具 | 输入依赖 |
| --- | --- | --- | --- |
| 1 | 教案DOCX | [教学设计定稿](../prompts/lesson-design.txt) | 教材与确认条件 |
| 2 | 学生工单及原版红字答案DOCX | [工单PROMPT](../prompts/worksheet.txt) | 教材与已完成并自检通过的教案 |
| 3 | 可编辑PPTX | [PPT定稿](../prompts/slides.txt) ＋ [模板规范](ppt-template.md) ＋ [通用起稿](../templates/vocational_teaching.pptx) ＋ [六页示例](../examples/magnetism-six-slides.pptx) | 教材、教案、学生工单及红字答案 |
| 4 | 汇总所有阶段的板书DOCX | [板书PROMPT](../prompts/board-plan.txt) | 全部前序成品，PPT已完成自动或等效检查、实际渲染与逐页看图 |
| 5 | 独立审查与一次修复 | [审查规则](independent-review.md) | 教材及全部成果实际文件与渲染预览图 |
| 可选 | SVG教学动画 | [动画分支](animation-workflow.md)与[SVG定稿](../prompts/animation.txt) | 教师明确选择、分镜确认 |

v0.5.4：沿用六类通用教学版式与磁学版式示例，保留详案与有效教学页门槛，加强主动思政挖掘和实际模板美学保真。教案五列表恢复[2,2,2,3,1]，学生活动约4.63cm；详案与每次连堂课至少12页有效教学内容按content-quality.md验收。来源与校验值见[文件清单](source-manifest.json)。

当次教师明确要求优先；SKILL.md确定范围及顺序，workflow.md确定阶段交接，提示词与内置脚本决定各成果内容和格式。教案提示词“不自动做PPT/工单”针对单独教案请求；整套模式明确要求这些成果。PPT整套模式明确交付采用通用教学版式制作的可编辑PPTX并保留所需制作记录。
工单不依赖未来PPT；PPT读取工单题目及答案；板书最后读取已完成课件。
六项确认和成果清单一次处理，不逐阶段要求教师审批。独立审查使用真实子代理，不用主代理自检替代。
最终成果目录仅交付DOCX与PPTX，核验用临时渲染图片存放在 `_render_check/` 供看图自检与审查，不混入交付物。
