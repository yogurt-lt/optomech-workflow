# 发布到 GitHub

账号 `yogurt-lt` 的具体上传及网页部署步骤见 [GitHub Pages 指南](GITHUB_PAGES.md)。当前版本已附静态导出器和自动部署工作流。

此目录是独立的发布候选，不要把原工作区整体初始化后推送。

1. 审阅 `examples/` 中的六套模型、处方、修改后的毛坯说明和公开范围。
2. 将 README、LICENSE 中的项目名和作者署名按需要调整。目前许可证草案为 MIT。
3. 在 GitHub 创建空仓库，例如 `optomech-workflow`，不附带自动初始化内容。
4. 在这个独立目录中执行以下命令；替换实际用户名与仓库地址。

```sh
git init
git add .
git diff --cached --stat
git commit -m "Initial release: six optical-to-mechanical demos and local workbench"
git branch -M main
git remote add origin https://github.com/yogurt-lt/optomech-workflow.git
git push -u origin main
```

在提交前检查暂存文件列表，只应有本候选目录内的通用代码与六个案例，不应有工作区其他设计文件。`.gitignore` 不是公开权限的替代品。

首个 Release 可标为 `v0.7.0-preview`，附上当前包、案例截面图、启动方式与验证范围。六个案例均已由用户在 SolidWorks 2023 运行检查；这不代表已经完成生产公差或应力验证。

后续修订时保留实际软件版本与验证记录。网页已提供三维预览、静态截面图及装配讲解。当前包同时提供本地 Python 工作台及 GitHub Pages 静态版发布工具；公开网站须在仓库创建、上传及部署成功后才可访问。
