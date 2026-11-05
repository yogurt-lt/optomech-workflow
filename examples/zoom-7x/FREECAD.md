# 变焦 v0.2

本版用于检查修改后的G3/G4结构，尚未更新公开网站。

- `Zoom_review_01_assembly.png`：光学装配剖面、点胶顺序、关键厚度及装入方向。
- `Zoom_review_02_motion.png`：完整YZ剖面及变焦驱动和位移曲线。
- SolidWorks2023：导入`Zoom_v02.bas`并保存宏，运行`main`；总装激活后运行`ZoomPosition1`至`ZoomPosition5`切换原始组态。运行产生独立时间戳目录，91个组件包含13片参考镜片。
- FreeCAD：打开`Zoom_v02.FCStd`，修改`ZoomControl.AngleDeg`（0～330°）并重算；需要重建时运行`Zoom_v02.FCMacro`。

本版修改5个Body，包括L8参考镜片。原始Zemax/STEP不变；L8新增Ø64机械余边，需要新制镜片，不能给现有镜片凭空增加材料。宏保留可编辑草图与特征；螺纹为名义接口，尚未画牙型。SolidWorks新宏由用户实际运行确认。
