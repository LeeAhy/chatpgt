# 部署到 Render

这是给“没有自己服务器”的情况准备的最省事方案。

## 你会得到什么

- 一个公网网址
- 不需要自己买服务器
- 你的网页仍然是“上传两个 Excel，直接下载结果”

## 你需要做的事

1. 注册一个 Render 账号
2. 把这个项目放到 GitHub 仓库
3. 在 Render 里新建一个 `Web Service`
4. 选择你的 GitHub 仓库
5. 使用下面的配置

## 配置

- Build Command: `pip install -r requirements.txt`
- Start Command: `python3 web_app.py`
- Plan: `Free`

## 重要说明

- Render 免费服务在一段时间没人访问后会休眠，再访问时会重新唤醒
- 免费服务的本地文件不会长期保存，不过这个项目是上传后立即处理并下载，所以不受影响

## 项目里已经准备好的文件

- [requirements.txt](/Users/chandelar/Documents/销售排单/requirements.txt)
- [render.yaml](/Users/chandelar/Documents/销售排单/render.yaml)
- [web_app.py](/Users/chandelar/Documents/销售排单/web_app.py)

