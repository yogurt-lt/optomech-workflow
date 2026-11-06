# v0.5 SolidWorks 优化宏

几何与原 v0.5 相同；保留原版 Projection_Zoom_v05.bas。本版不更新网站。

## 运行

将 Projection_Zoom_v05_Fast.bas 导入新宏项目，保存为 swp。第一次选择 main：完整生成，默认每32次切除做一次显式重建检查，各零件末尾仍完整检查并保存。

后续选择 mainReuse：选择上一次本优化版产生的 Optomech_时间 文件夹。仅复用几何键、保存文件大小/修改时间、模型属性与单实体检查都匹配的零件，复制到新的输出目录后重新装配。原目录不修改。没有匹配或文件仍在打开时，该零件重新生成；旧版宏生成、没有 .reuse 文件的零件不能自动复用。保留每个 SLDPRT 旁的 .reuse 记录；它用于工程缓存有效性检查，不是防篡改的密码学认证。

如需逐次定位异常，运行 mainDiagnostic：恢复每次切除后检查，仍使用新的草图查找和显示优化。

## 实现

- 优先用 GetLastFeatureAdded 取得刚创建的3D草图，校验类型和名称；不符合时回退原特征树查找。
- 草图创建暂不逐条显示，模型视图暂停中间刷新；成功或错误退出时恢复原显示设置。
- 显式切槽重建由1004次减到28个分批检查点，不改变原始几何、切除顺序、截面数据及装配变换；每次 API 切除失败仍立即停止，每个零件末尾仍检查单实体及轴向范围。
- 保留短段切槽和小 VBA 过程修复；未在这一版改成尚未验证的长连续槽或齿轮阵列。
- build_log.txt 中 PART seconds 是各零件生成/复用耗时；CUT checkpoint 是该零件累计切槽检查点及耗时；ASSEMBLY seconds 和 TOTAL seconds 为装配及总耗时。跨午夜计时按一天回绕处理。

本版静态检查通过，尚未在 SolidWorks 实际编译运行或计时；不承诺具体提速倍数。首轮仍有1004个完整几何切除特征，不能保证立刻完成。遇到错误保留日志；若未完成某个零件，下次 mainReuse 可复用此前已成功保存的零件。

## API依据

- [GetLastFeatureAdded](https://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~GetLastFeatureAdded.html)
- [草图集中创建与延迟显示](https://help.solidworks.com/2023/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager~DisplayWhenAdded.html)
- [模型视图刷新控制](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.Interop.sldworks.IModelView~EnableGraphicsUpdate.html)
