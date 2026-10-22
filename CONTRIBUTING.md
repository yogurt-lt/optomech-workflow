# Contributing

欢迎提交可复现案例、几何检查改进、文档或 SolidWorks 兼容性反馈。

请先确认案例可以公开，不要附带客户或公司的非公开设计、账号信息、绝对本机路径或软件安装包。问题中建议包含输入文件、方案 JSON、SolidWorks 版本、错误截图和宏输出日志。发布完整处方前请自行确认数据的公开权限。

修改几何算法后运行 `python -m unittest discover -s tests -v`，并对两个示例执行 `python -m optomech check ...`。涉及 SolidWorks API 的修改请说明是否实机运行，以及验证的软件版本。不要把源代码静态检查写成 SolidWorks 实机通过。

默认不提交 `.SLDPRT`、`.SLDASM`、`.ZDA`、自动生成日志和临时文件。需要分享 CAD 结果时可放到经过审阅的 Release 附件。
