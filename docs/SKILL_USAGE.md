# 光机结构设计 Skill

[English](SKILL_USAGE.en.md)

下载 `optomech-design-skill.zip`，解压得到 `optomech-design/`。在支持 `SKILL.md` 的 AI 助手中，将整个目录加入其技能目录，按该工具的方式加载或调用 `optomech-design`。保留 `references/`、`scripts/` 和 `assets/`，不要只复制入口文件。若工具不支持自动发现，可直接让助手读取 `SKILL.md` 并按其流程执行。

建议首次使用时提供：可公开或有权处理的 Zemax / STEP 文件、单位与光轴、CAD 软件与版本（SolidWorks 或 FreeCAD）、是否允许修改镜片机械外缘、是否需要通光检查，以及大致制造方式。无需上传与本次工作无关的产品文件。

可直接使用的提示词：

> 使用 optomech-design。先检查这个文件夹里的 Zemax 和 STEP 是否一致，再查阅适合该镜头类型的装配资料。先解释镜筒分段、装入方向、沉靠、固定方式和传力路径，再生成可编辑的镜筒、隔圈、压环宏及同一版本的截面图。不要只按口径排序自动决定方案；标明哪些检查已经执行。请用中文交流。

Skill 包含我们从六个 demo 中提炼的设计规则、SolidWorks 2023 的常见错误修复经验，以及独立的 Python 文件分析、截面校核和基础宏生成工具。Python 3.10+ 即可使用这些脚本；执行 CAD 宏需要本机安装所选 CAD；本轮已在 SolidWorks 2023 和 FreeCAD 1.1.4 中逐案例运行检查。

脚本支持明确方案的数值检查；FreeCAD 后端还支持明确声明的径向孔、开口窗口和扳手槽。SolidWorks 通用生成器仍以旋转主体为主，完整孔槽使用已验证案例宏或另行实现；装配方案由助手结合文件和资料推理。Skill 不会自动将任意 STEP 变成已完成生产验证的设计。

**首次试用命令**（进入解压目录）：

```sh
python scripts/workflow.py analyze --step lens.STP --zmx lens.zmx
python scripts/workflow.py check design_plan.json --report checks.json
python scripts/workflow.py macro design_plan.json --output Concept.bas --include-glass
python scripts/workflow.py macro design_plan.json --cad freecad --output Concept.FCMacro --include-glass
```

英文版工具指南在 `references/plan-and-tools.md`。图形检查通过、宏成功运行和完成真实制造验证分别记录。聊天里提出的尺寸属于各个案例，不会被设定为所有镜头的通用加工限制。

新版 Skill 将设计流程与 CAD 后端分离。FreeCAD 的原生特征适配见 `references/freecad.md`；若方案包含未迁移的孔槽，默认会停止，只有明确选择主体 baseline（`--allow-base-only`）才继续，并列出遗漏。SolidWorks 宏不能直接转换成 FreeCAD 宏；应从同一方案重新生成。六个 demo 已分别在 FreeCAD 运行并经用户确认；新方案和修改版仍需各自验证。

分发包包含通用代码和提炼后的经验，不包含私人对话、公司镜头文件或账号凭据。六个公开示例在完整项目包中提供，独立 Skill 不依赖这些输入。

## 调焦机构

Skill 包含整组手动与单组电动对焦推理参考 `references/focusing-and-apertures.md`，以及几何行程 CSV 工具 `scripts/motion_law.py`。先确定运动组件、固定基准和装配路径，再画凸轮、滚子或螺旋副。F35 示例的 ±2 mm 行程与尺寸不是通用默认；光阑表面遮光方案仍需光学校核。
