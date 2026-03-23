# WEB_edit
启动需输入：`python3 /Users/anita/code/serve_site.py`
本地打开：`http://127.0.0.1:8000/`
公开网站链接：https://anitahy2000-hub.github.io/WEB_edit/
以后你每次改完网页，只要执行：`bash /Users/anita/code/publish_site.sh`

## 在线后端部署

这个仓库已经补好了 Render 部署文件：

- [render.yaml](/Users/anita/code/render.yaml)
- [requirements.txt](/Users/anita/code/requirements.txt)

部署步骤：

1. 把当前仓库推到 GitHub。
2. 打开 Render，新建 `Blueprint`。
3. 选择这个 GitHub 仓库。
4. Render 会自动识别 [render.yaml](/Users/anita/code/render.yaml) 并创建 Python Web Service。
5. 部署完成后，访问 Render 分配的域名即可在线使用 `公众号导入` 和 `DOCX 导出`。

说明：

- GitHub Pages 版本仍然只是静态页面。
- 真正可在线使用的后端，请访问 Render 上的服务地址。
- 线上服务根路径 `/` 会直接打开 `text_format_tool.html`。
