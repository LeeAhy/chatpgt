# Netlify Blobs 持久保存配置

这个方案让 Render 免费后端继续处理 Excel，同时把网站状态和当前最新版文件保存到 Netlify Blobs。

## 需要保存到 Netlify Blobs 的内容

- `current_sales.xlsx`：当前共用销售排单
- `latest_generated.xlsx`：当前最新版下载文件
- `metadata.json`：每位业务上传状态、上传时间、最近回填记录

## 1. 部署 Netlify

把本项目上传到 GitHub 后，在 Netlify 新建站点。

Netlify 构建设置：

- Build command: `npm install`
- Publish directory: `netlify-site`
- Functions directory: `netlify/functions`

部署完成后，Netlify 会生成一个公开网址，例如：

```text
https://your-site.netlify.app
```

存储接口地址就是：

```text
https://your-site.netlify.app/api/blob-storage
```

## 2. 在 Netlify 设置密钥

进入 Netlify 项目：

```text
Site configuration -> Environment variables
```

新增：

```text
SALES_STORAGE_SECRET=自己设置一串长密码
```

例如：

```text
SALES_STORAGE_SECRET=sales-2026-private-storage-key
```

保存后重新部署 Netlify。

## 3. 在 Render 设置同样的环境变量

进入 Render 的 `sales-upload` 服务：

```text
Environment
```

新增两个变量：

```text
NETLIFY_BLOBS_ENDPOINT=https://your-site.netlify.app/api/blob-storage
SALES_STORAGE_SECRET=和 Netlify 完全一样的密钥
```

保存后 Render 会重新部署。

## 4. 验证是否生效

打开：

```text
https://sales-upload.onrender.com/status
```

如果看到：

```text
当前已启用 Netlify Blobs 远程保存
```

就说明已经启用。

之后你上传本周共用销售排单、每位业务预测、在线编辑当前表格，网站都会自动同步到 Netlify Blobs。

## 注意

- Netlify Blobs 负责保存数据，Render 仍负责处理 Excel。
- 如果 Render 免费实例休眠，首次打开仍可能慢几十秒，但数据不会因为 Render 重启而丢。
- 如果重新部署 Render，网站会自动从 Netlify Blobs 拉回当前共用排单和业务状态。
