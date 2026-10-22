# 光机结构设计 Skill

[English](SKILL_USAGE.en.md)

下载 `optomech-design-skill.zip`，解压得到 `optomech-design/`。在支持 `SKILL.md` 的 AI 助手中，将整个目录加入其技能目录，按该工具的方式加载或调用 `optomech-design`。保留 `references/`、`scripts/` 和 `assets/`，不要只复制入口文件。若工具不支持自动发现，可直接让助手读取 `SKILL.md` 并按其流程执行。

建议首次使用时提供：可公开或有权处理的 Zemax / STEP 文件、单位与光轴、SolidWorks 版本、是否允许修改镜片机械外缘、是否需要通光检查，以及大致制造方式。无需上传与本次工作无关的产品文件。

可直接使用的提示词：

> 使用 optomech-design。先检查这个文件夹里的 Zemax 和 STEP 是否一致，再查阅适合该镜头类型的装配资料。先解释镜筒分段、装入方向、沉靠、固定方式和传力路径，再生成可编辑的镜筒、隔圈、压环宏及同一版本的截面图。不要只按口径排序自动决定方案；标明哪些检查已经执行。请用中文交流。

Skill 包含我们从六个 demo 中提炼的设计规则、SolidWorks 2023 的常见错误修复经验，以及独立的 Python 文件分析、截面校核和基础宏生成工具。Python 3.10+ 即可使用这些脚本；执行 CAD 宏仍需要本机 SolidWorks。

脚本支持的是明确方案的数值检查和旋转主体生成。专用孔槽必须另外实现并验证；装配方案由助手结合文件和资料推理。Skill 不会自动将任意 STEP 变成已完成生产验证的设计。

**首次试用命令**（进入解压目录）：

```sh
python scripts/workflow.py analyze --step lens.STP --zmx lens.zmx
python scripts/workflow.py check design_plan.json --report checks.json
python scripts/workflow.py macro design_plan.json --output Concept.bas --include-glass
```

英文版工具指南在 `references/plan-and-tools.md`。图形检查通过、宏成功运行和完成真实制造验证分别记录。聊天里提出的尺寸属于各个案例，不会被设定为所有镜头的通用加工限制。

分发包包含通用代码和提炼后的经验，不包含私人对话、公司镜头文件或账号凭据。六个公开示例在完整项目包中提供，独立 Skill 不依赖这些输入。
