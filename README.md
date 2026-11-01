# 整堂课资源制作 v0.5.1

面向中职教师的教学资源工作流：教案DOCX → 学生工单与红字答案DOCX → 可编辑PPTX → 板书DOCX → 独立审查。可选HTML/SVG动画须明确选择。

## 安装
下载 `lesson-resource-workflow-v0.5.1-public.zip` 原始附件。在WorkBuddy的“专家·技能·连接器 → 技能 → 添加技能 → 上传技能”选择该ZIP。不要用整个GitHub仓库的源码ZIP替代安装附件。其他宿主可将解压后的lesson-resource-workflow目录放入其技能目录。

## 启动示例
“使用整堂课资源制作技能，根据我提供的教材制作本次资源。汽车运用与维修专业，学生已学基础电路，普通教室理论课，有黑板、投影与打印条件，用于普通新授课；2学时，每学时40分钟。请先回显需要确认的条件，再连续制作。”
实际课题与教材另行提供；示例条件不能自动套用到真实班级。

## 环境与依赖
需宿主支持读取教材、生成可编辑DOCX/PPTX、运行检查和查看实际预览。完整独立审查还需真实子代理能力。Python检查依赖安装：`python -m pip install -r <技能目录>/scripts/requirements.txt`。渲染另需LibreOffice以及PyMuPDF或Poppler；可用`python -m pip install PyMuPDF`安装前者。由用户按本机环境安装，不要求技能自动修改系统。
示例：`python <技能目录>/scripts/check_deck.py <课件.pptx> --render-dir <临时目录>`；路径含空格时加引号。模板维护检查才使用`--template`。退出0仅代表自动检查与渲染完成，还须看图；1表示错误；2表示渲染未完成。
缺少依赖时可使用宿主等效工具，逐项说明检查范围。无子代理不能宣称独立审查通过。

## 使用边界
40分钟一学时、80分钟连堂课和40行表格是可调整的默认约定，不是全国统一要求。按本校实际时长确认。六页示例不是完整课程；专业内容必须由授课教师核验。许可见LICENSE；素材说明见技能内references/attribution.md。

## 验证状态
本机脚本及渲染结果见发布复查报告。WorkBuddy市场必填元数据已补齐；本版本尚未完成WorkBuddy桌面实际导入和整套生成试跑，按公开测试版发布，不宣称全流程实测通过。平台是否接受与上架以实际导入及审核为准。

## GitHub发布
提交本目录README、LICENSE、CHANGELOG及技能源码；将安装ZIP与SHA256SUMS.txt附在v0.5.1 Release，标记Pre-release。不要提交本地审查缓存、教材扫描件或个人资料。
