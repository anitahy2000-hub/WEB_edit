# GitHub Pages 部署说明

这个项目已经整理成可直接发布到 GitHub Pages 的版本，入口文件是：

- `docs/index.html`

## 当前能力

- 可直接在线打开网页
- 可上传 `.docx` 并在浏览器里解析为图文内容
- 可继续编辑标题、加粗、图片说明
- 可复制 Markdown、HTML
- 可在浏览器里导出 `.docx`

## 部署步骤

1. 在 GitHub 上新建一个仓库。
2. 把当前目录内容上传到仓库根目录。
3. 默认分支使用 `main`。
4. 推送后，GitHub Actions 会自动运行：
   - 工作流文件在 `.github/workflows/deploy-pages.yml`
5. 在仓库设置中打开：
   - `Settings` -> `Pages`
   - `Build and deployment` 选择 `GitHub Actions`
6. 等待工作流完成后，就会得到公开访问链接。

## 说明

- GitHub Pages 最终发布的是 `docs/` 目录内容。
- 线上入口页是 `docs/index.html`。
- 如果你后续继续修改 `text_format_tool.html`，记得同步更新 `docs/index.html`。
