# 服务器部署说明（域名访问，无需用户本地后端）

本文档目标：将本项目部署到服务器后，后端持续运行，用户只通过域名访问即可登录与使用。

## 1. 准备目录

```bash
sudo mkdir -p /opt/time-test1
sudo chown -R $USER:$USER /opt/time-test1
```

将项目文件上传到 `/opt/time-test1`。

## 2. 安装 Python 依赖

```bash
cd /opt/time-test1
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 3. 配置管理员账号（可增删，无需重启）

编辑：

`/opt/time-test1/data/admin_users.json`

示例：

```json
{
  "users": [
    { "username": "admin", "password": "admin123" },
    { "username": "teacherA", "password": "abc123456" }
  ]
}
```

支持 `password_sha256` 字段（更安全）：

```json
{
  "username": "secureUser",
  "password_sha256": "你的sha256十六进制字符串"
}
```

## 4. systemd 常驻后端

复制模板并修改路径/用户：

```bash
sudo cp /opt/time-test1/deploy/volunteer-api.service /etc/systemd/system/volunteer-api.service
sudo nano /etc/systemd/system/volunteer-api.service
```

至少确认这几项：

- `User`
- `Group`
- `WorkingDirectory`
- `ExecStart`

启动并设置开机自启：

```bash
sudo systemctl daemon-reload
sudo systemctl enable volunteer-api
sudo systemctl restart volunteer-api
sudo systemctl status volunteer-api
```

## 5. Nginx 配置域名

复制模板并修改：

```bash
sudo cp /opt/time-test1/deploy/nginx.conf /etc/nginx/sites-available/time-test1.conf
sudo nano /etc/nginx/sites-available/time-test1.conf
```

至少修改：

- `server_name`
- `root`

启用站点：

```bash
sudo ln -sf /etc/nginx/sites-available/time-test1.conf /etc/nginx/sites-enabled/time-test1.conf
sudo nginx -t
sudo systemctl reload nginx
```

## 6. 验证

```bash
curl -i http://127.0.0.1:8000/health
curl -i http://your-domain.com/health
```

浏览器打开：

- `http://your-domain.com/`

如果页面能打开但登录提示无法连接后端，通常是 `/api` 反向代理未生效，请检查 Nginx 配置与重载状态。

