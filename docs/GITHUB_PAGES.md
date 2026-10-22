# yogurt-lt：上传项目并启用网页

项目建议使用独立仓库 `yogurt-lt/optomech-workflow`，与个人主页仓库 `yogurt-lt.github.io` 分开管理。

## 首次上传

1. 在 GitHub 新建名为 `optomech-workflow` 的 Public 仓库。保留项目自带 README 和 LICENSE，创建时无需另选。
2. 用 GitHub Desktop 将新仓库克隆到一个独立目录。
3. 解压发布包，将 `optomech-workflow/` **里面的全部内容**复制到克隆目录，包括 `.github/` 和 `.gitignore`。仓库根目录应直接看到 `README.md`、`server.py`、`tools/`、`web/`、`examples/`、`skills/`。
4. 在 GitHub Desktop 填写说明并执行 Commit to main、Push origin。
5. 打开仓库 **Settings → Pages → Build and deployment → Source → GitHub Actions**。
6. 在 Actions 中打开 `Deploy bilingual gallery to GitHub Pages`。若首次推送时尚未启用 Pages 导致失败，启用后选择 Re-run all jobs；也可使用 Run workflow 手动运行。

部署成功后的预期地址：

`https://yogurt-lt.github.io/optomech-workflow/`

当前交付的是发布候选文件，不表示 GitHub 仓库已创建或网页已经上线。最终 URL 以成功的 Actions 部署结果为准。

## 更新与功能

以后修改案例、Skill 或页面后，Commit、Push；工作流会重新生成并部署。静态网页保留六案例选择、三维预览、中文/English、截面图、装配讲解以及宏/Skill 下载。各案例的既有宏保持不变。

文件上传分析与修改版宏生成使用本地 `python server.py` 工作台；Pages 版本展示已检查案例，不依赖 Python 后端。

本地检查静态导出：

```sh
python tools/build_pages.py
python -m http.server 8766 --bind 127.0.0.1 --directory dist
```

然后打开 `http://127.0.0.1:8766/`。静态版所有资源使用相对路径，也支持项目子路径部署。

只上传独立发布候选目录；不要上传上级 `ai机械设计` 工作区。该工作区含未公开的公司设计。

## English

Create the public repository `yogurt-lt/optomech-workflow`, copy the package contents to its root (including `.github/`), commit and push to `main`. In **Settings → Pages**, select **GitHub Actions** as the source. Re-run the included deployment workflow if the first run occurred before Pages was enabled.

The expected URL after successful deployment is `https://yogurt-lt.github.io/optomech-workflow/`. The static gallery supports the reviewed demos, bilingual UI/drawings, 3D preview and downloads. The Python workbench remains available locally for file analysis and custom macro generation. These files prepare deployment; they do not by themselves create a repository or publish the site.
