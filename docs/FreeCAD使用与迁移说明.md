# FreeCAD 使用与迁移 v0.7

六个案例已由用户分别运行检查确认，FreeCAD 原生验证版本为 1.1.4。每个 `examples/<case>/` 提供 `FreeCAD.FCMacro`、`FreeCAD_Model.FCStd`、中英文指南和 `freecad_report.json`。

| 案例 | 机械件 | 参考玻璃 |
|---|---:|---:|
| 定焦投影 PF-2，三个可拆胶接短框 | 9 | 10 |
| 胶片投影 | 7 | 5 |
| 显微物镜 | 7 | 10 |
| 手机三台阶 | 5 | 6 |
| 单反加厚 v5 | 13 | 11 |
| 广角 | 8 | 10 |

## 运行

从 [FreeCAD 官方下载页](https://www.freecad.org/downloads.php?lang=en) 安装 Windows 稳定版。打开软件，菜单“宏 → 宏”，设置用户宏位置到案例目录，选中 `FreeCAD.FCMacro` 并执行。运行宏不需要另外安装 Python；生成新宏的命令行工具需要 Python 3.10+。

选择输出目录后，新建带时间戳的文件夹，保存原生 `.FCStd` 和 `build_report.json`。也可以直接打开已提供的 `FreeCAD_Model.FCStd`。模型树中每个零件独立 Body，保留 Sketch、Revolution 和适用的 Pocket。隐藏外筒查看内部，或“视图 → 剖切视图”勾选 X、偏移 0，查看沿光轴剖面。显示剖切不修改实体，切口未填色不代表空心错误。

## 支持与边界

FreeCAD 后端支持明确指定的径向孔、开口点胶窗口、两个相对扳手槽及原生保存。宏、模型与实际报告是逐案例记录；用户确认不等同于通光、应力、公差和制造验证。胶、真实螺纹、螺钉和 Assembly 工作台关节未生成。部分曲面接触存在采样参考体微小交叠，按案例报告保留；不掩盖为“全部精确无干涉”。

通用 CLI：

```sh
python skills/optomech-design/scripts/workflow.py macro examples/phone/plan.json --cad freecad --output Phone.FCMacro --include-glass
```

从同一方案生成不同软件的原生特征，而不是翻译 VBA。静态源检查不能代替目标 CAD 执行；报告中应确认 success 和预期零件数，而不只是看到第一个筒体。六个案例对应不同装入方向，具体看每个案例的 FREECAD.md。
