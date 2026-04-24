# 志愿时长查询与补录系统

一个面向学生与学院/办公室管理员使用的志愿时长管理项目，当前采用 `静态单页前端 + FastAPI + Excel 数据源` 的实现方式，适合快速上线查询、公示、补录与后台处理流程。

系统主要解决这几类场景：

- 学生查询各年级志愿时长与项数。
- 2022 级学生进行毕业达标校验。
- 学生提交校内/校外时长补录材料。
- 管理员下载总表、处理补录记录、批量录入活动时长。
- 高级管理员上传新的公示总表。

当前前端入口是根目录的 `index.html`，后端入口是 `api.py`，主要业务数据保存在 `data/` 下的 Excel 文件中。

这份 README 默认面向“使用 vibe coding 接手项目的人”编写，也就是说：

- 接手者不一定会长期手写代码，但会频繁借助 AI/Codex 一类工具理解、修改和排查项目。
- 文档重点不是完整的传统架构设计，而是帮助接手者快速建立“这个项目现在到底怎么运行、数据放哪、哪些地方不能乱改”的上下文。
- 如果你准备让 AI 帮你改项目，先读完本文里的“项目结构”“数据文件与运行约定”“鉴权与权限现状”“已知限制 / 风险提醒”四节，再开始动手。

## 系统角色与能力

### 学生端

- 时长查询：按 `姓名 + 学号 + 年级` 查询汇总结果。
- 2022 级毕业达标查询：校验总时长是否达到 `32h`、项数是否达到 `8`。
- 原始总表下载：允许学生下载 2022/2023 级原始总表文件。
- 补录提交：
  - 校内补录：上传活动证明图片。
  - 校外补录：上传证明图片、活动照片、活动描述与项数。

### 管理员端

- 登录后台查看与处理管理功能。
- 下载各年级已上传的最新总表。
- 批量录入活动时长到总表新列。
- 查看、批量更新、批量删除补录记录。
- 导出 `学生时长补录汇总.xlsx`。
- 快捷合并多分表 Excel 为单一汇总表。

### 高级管理员

- 通过二次验证后，上传新的公示总表。
- 上传后系统会自动合并分表并生成 CSV，用于更快的查询。

## 项目结构

```text
time-test1/
├── api.py                      # FastAPI 入口，接口、鉴权、上传下载逻辑
├── config.py                   # 路径、上传限制、数据文件位置配置
├── index.html                  # 单页前端，包含学生端与管理员端界面
├── requirements.txt            # Python 依赖
├── services/
│   └── excel_service.py        # Excel 查询、合并、追加写入、导出逻辑
├── data/                       # 总表、补录汇总、反馈汇总、管理员账号
├── uploads/
│   └── proofs/                 # 学生上传的证明文件
├── deploy/                     # Nginx / systemd 部署模板
└── DEPLOYMENT.md               # 详细部署说明
```

重点目录说明：

- `api.py`：定义学生端、管理员端和高级管理员相关接口，并处理上传、下载、鉴权和健康检查。
- `services/excel_service.py`：封装所有 Excel 读写逻辑，包括总表合并、CSV 生成、补录汇总维护、批量更新等。
- `config.py`：统一管理 `data/`、`uploads/` 等目录路径，以及上传大小和扩展名限制。
- `index.html`：静态页面，直接在浏览器打开即可使用；包含学生查询、补录、管理员后台等交互。
- `deploy/`：保存生产环境使用的 `nginx.conf` 和 `systemd` 服务模板。
- `data/`：保存按年级区分的总表、反馈汇总、补录汇总和管理员账号文件。
- `uploads/proofs/`：保存学生补录或反馈时上传的证明文件与照片。

## 面向 Vibe Coding 接手的阅读顺序

如果你的后续维护方式是“先让 AI 帮你读仓库，再逐步改功能”，建议按这个顺序建立上下文：

1. 先看 `README.md`
   先搞清楚系统角色、数据落点、部署方式和风险边界。
2. 再看 `index.html`
   前端是单文件页面，学生端、管理员端、按钮行为和请求入口都集中在这里。
3. 再看 `api.py`
   这里能快速看出全部对外接口、鉴权方式、上传下载逻辑、哪些接口给学生用、哪些给管理员用。
4. 最后看 `services/excel_service.py`
   这里是最容易被 AI 改坏、但也是最关键的业务逻辑层，尤其是 Excel 合并、列名兼容、补录汇总追加、按学号/姓名聚合等逻辑。

如果你让 AI 帮你改功能，建议直接把下面这些事实先告诉它：

- 这是 `静态 index.html + FastAPI + Excel 文件持久化` 项目，不是前后端分仓、也不是数据库项目。
- `data/` 和 `uploads/` 中可能有真实业务数据，不能随便删除、覆盖或批量清空。
- `api.py` 中有管理员和高级管理员两层权限，且 token 是内存态。
- 查询速度依赖 `xlsx -> csv` 的转换结果，不要随手删掉生成的 CSV 逻辑。
- `services/excel_service.py` 里有很多对脏 Excel、别名列名、多分表的兼容代码，改动前必须先理解再下手。

## 技术栈

- 后端：FastAPI
- 前端：原生 HTML + Bootstrap 5
- 数据处理：pandas、openpyxl
- 文件上传：python-multipart
- 部署：Nginx + systemd + uvicorn

## 本地运行

### 1. 安装依赖

建议使用 Python 3.11+ 或与现有环境一致的 Python 3 版本。

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. 启动后端

推荐开发启动命令：

```bash
uvicorn api:app --host 127.0.0.1 --port 8000 --reload
```

也可以直接运行：

```bash
python api.py
```

### 3. 打开前端

直接在浏览器中打开根目录下的 `index.html` 即可。

前端内置了如下 API 基址策略：

- 如果页面运行在 `file://` 下，默认请求 `http://127.0.0.1:8000`
- 如果页面部署在站点中，则默认走当前站点同源地址

### 4. 健康检查

后端启动后可通过以下接口确认服务是否正常：

```bash
curl http://127.0.0.1:8000/health
```

预期返回：

```json
{"status":"ok"}
```

## 用 AI / Codex 维护时的建议

如果接手者主要通过 vibe coding 维护这个项目，建议遵守下面这些原则：

- 先让 AI 做“阅读和总结”，再让它改代码，不要一上来就让它重构。
- 每次改动只针对一个目标，例如“修复补录上传”“增加一个接口字段”“调整管理员页面文案”，避免一次让 AI 同时改前端、后端、Excel 逻辑和部署。
- 只要改到 `services/excel_service.py`，就默认这是高风险改动，因为这里直接影响真实表格数据。
- 只要改到上传、删除、批量更新逻辑，就要先确认不会误删 `data/`、`uploads/` 中的真实文件。
- 不要让 AI 擅自“清理仓库”或“删除看起来没用的文件”，因为当前目录里混有真实运行数据、证书、部署文件和历史材料。
- 改完后至少手工验证一遍：查询是否正常、后台是否还能登录、上传/下载是否还能走通。

更具体一点，适合直接交给 AI 的任务类型包括：

- 补充或修改 README、部署说明、注释。
- 调整 `index.html` 中的前端文案、表单提示、按钮流程。
- 新增低风险接口字段或返回提示。
- 修复某个已知接口报错，并保持原有数据格式不变。

不适合直接让 AI 大范围自由发挥的任务包括：

- 重构整个 Excel 处理层。
- 一次性替换鉴权机制、部署方式和数据结构。
- 在未备份数据前批量改写 `data/` 下文件。
- 在没有理解现有上传路径约定前重做附件存储逻辑。

## 数据文件与运行约定

系统当前以 Excel 作为主要持久化方式，而不是数据库。

### 总表文件

按年级维护：

- `data/学生志愿时长总表_2022.xlsx`
- `data/学生志愿时长总表_2023.xlsx`
- `data/学生志愿时长总表_2024.xlsx`
- `data/学生志愿时长总表_2025.xlsx`

兼容旧逻辑的默认文件：

- `data/学生志愿时长总表.xlsx`

上传总表后，系统会额外生成同名 CSV 文件，例如：

- `data/学生志愿时长总表_2024.csv`

这样查询接口优先读 CSV，从而提升读取速度。

### 汇总文件

- `data/学生反馈汇总.xlsx`：记录学生反馈与证明文件相对路径。
- `data/学生时长补录汇总.xlsx`：记录校内/校外补录申请、处理状态、管理员备注等。
- `data/admin_users.json`：管理员账号配置文件。

### 上传文件

上传的证明文件保存在：

- `uploads/proofs/`
- `uploads/proofs/topup_internal/`
- `uploads/proofs/topup_external/`

系统在汇总表中保存的是相对路径，管理员通过接口进行预览或下载，而不是直接把二进制内容写入 Excel。

### 管理员账号

`data/admin_users.json` 支持两种密码格式：

- `password`：明文密码，仅建议本地临时使用。
- `password_sha256`：SHA-256 哈希，生产环境更推荐使用。

当前项目里已经存在的普通管理员账号如下：

```json
{
  "users": [
    { "username": "lipangbo", "password": "123" },
    { "username": "marius", "password": "456" },
    { "username": "office001", "password": "office234" },
    { "username": "admin123", "password": "admin456" }
  ]
}
```

另外，后端里还存在一组单独写死在 `api.py` 中的高级管理员验证账号：

```text
高级账号：admin123
高级密码：admin456
```

也就是说，接手时需要区分两种权限：

- 普通管理员账号：来自 `data/admin_users.json`
- 高级管理员账号：当前写死在 `api.py` 的 `/api/admin/senior-login` 逻辑中

如果后续继续沿用这些账号，至少要先确认是否需要改密码；如果准备正式交接，建议尽快替换为新的账号体系。

如果要改成哈希存储，可以参考下面的格式：

```json
{
  "users": [
    {
      "username": "admin",
      "password_sha256": "your_sha256_hash"
    }
  ]
}
```

## 核心接口一览

README 这里只列用途和场景，不展开成完整 API 文档。

### 学生侧接口

- `GET /api/search`
  - 用途：按姓名、学号、年级查询学生志愿时长。
- `GET /api/master-last-upload`
  - 用途：查询指定年级总表最后上传时间。
- `GET /api/public/master-file-status`
  - 用途：公开查看某年级总表是否已上传及最后更新时间。
- `GET /api/download-master-raw`
  - 用途：学生下载 2022/2023 级原始总表。
- `POST /api/feedback`
  - 用途：提交问题反馈与证明文件。
- `POST /api/topup/internal`
  - 用途：提交校内时长补录。
- `POST /api/topup/external`
  - 用途：提交校外时长补录。

### 管理员接口

- `POST /api/admin/login`
  - 用途：普通管理员登录。
- `POST /api/admin/senior-login`
  - 用途：高级管理员二次验证。
- `POST /api/admin/upload-master`
  - 用途：上传某年级公示总表。
- `POST /api/admin/append-durations`
  - 用途：将待录入活动时长追加到总表新列。
- `POST /api/admin/merge-grades`
  - 用途：合并多分表 Excel 并返回合并结果。
- `GET /api/admin/master-files`
  - 用途：查看各年级总表状态。
- `GET /api/admin/download-master`
  - 用途：下载指定年级总表或原始文件。
- `GET /api/admin/download-feedback`
  - 用途：导出学生反馈汇总。
- `GET /api/admin/download-topup`
  - 用途：导出补录汇总。
- `GET /api/admin/topup-records`
  - 用途：查看补录记录。
- `POST /api/admin/delete-topup`
  - 用途：删除单条补录记录。
- `POST /api/admin/delete-topup-batch`
  - 用途：批量删除补录记录。
- `POST /api/admin/update-topup-batch`
  - 用途：批量更新补录状态或管理员备注。
- `GET /api/admin/uploaded-file`
  - 用途：管理员预览上传的附件。

### 配置入口

- `config.BaseConfig`
  - 用途：统一管理数据目录、上传目录、文件名规则和上传限制。

## 鉴权与权限现状

当前实现是偏轻量的内部工具式方案，README 如实记录以下现状：

- 普通管理员通过 `POST /api/admin/login` 登录。
- 高级管理员通过 `POST /api/admin/senior-login` 进行二次验证。
- 普通管理员 token 和高级管理员 token 都保存在后端内存中。
- token 默认有效期为 2 小时。
- 服务重启后，内存中的 token 会全部失效，需要重新登录。

这意味着当前方案更适合小规模、低并发、少量管理员的使用场景，不适合作为严格意义上的企业级权限系统。

## 部署方式

生产环境当前采用如下结构：

```text
Browser
  -> Nginx
  -> FastAPI / uvicorn (127.0.0.1:8000)
  -> data/*.xlsx + uploads/proofs/*
```

部署相关文件：

- [DEPLOYMENT.md](/Users/doormi/Documents/time-test1/DEPLOYMENT.md)
- [deploy/nginx.conf](/Users/doormi/Documents/time-test1/deploy/nginx.conf)
- [deploy/volunteer-api.service](/Users/doormi/Documents/time-test1/deploy/volunteer-api.service)
- [deploy/time-test1.service](/Users/doormi/Documents/time-test1/deploy/time-test1.service)

README 只保留摘要，详细步骤请查看 `DEPLOYMENT.md`。

接手服务器时，还应该同时阅读下面这份交接文档：

- [SERVER_HANDOVER.md](/Users/doormi/Documents/time-test1/SERVER_HANDOVER.md)

### 简版部署流程

1. 将项目上传到服务器目录。
2. 创建 Python 虚拟环境并安装依赖。
3. 按需配置 `data/admin_users.json`。
4. 使用 `systemd` 启动 `uvicorn api:app`。
5. 使用 Nginx 代理 `/api/` 与 `/health`，并将 `/` 指向 `index.html`。

说明：

- 当前仓库内 `deploy/nginx.conf` 已体现线上站点的现有反向代理方式。
- README 不把现有域名、服务器地址或私密配置当作默认模板推荐给新环境。
- 新环境部署时，应根据自己的域名、证书路径、用户和目录重新修改配置。
- 当前线上服务器是上一任维护者个人购买的，不能默认继续使用；接手方需要自备服务器、域名和证书后重新部署。

## 开发与维护建议

### 典型维护动作

- 更换总表：通过高级管理员入口上传新的年级总表。
- 查看补录：管理员后台进入补录记录处理页面。
- 导出补录汇总：管理员后台下载 `学生时长补录汇总.xlsx`。
- 调整管理员账号：直接修改 `data/admin_users.json`，当前实现支持动态读取，无需重启服务。

### 适合接手前先确认的事项

- `data/` 中各年级总表是否为当前学年最新版。
- `uploads/proofs/` 是否需要做备份或清理。
- `admin_users.json` 是否仍存在明文密码。
- 部署脚本和证书路径是否仍与服务器现状一致。
- 如果主要通过 AI 维护，是否已经先把当前目标、涉及文件和不能破坏的数据边界讲清楚。

## 已知限制 / 风险提醒

当前实现能满足实际业务流转，但仍有一些需要明确的限制：

- 不是数据库方案，数据一致性、事务能力和并发处理能力有限。
- Excel 作为主存储时，更适合中小规模内部使用，不适合高并发场景。
- 管理员 token 与高级管理员 token 均为内存态，服务重启后会失效。
- 当前高级管理员验证逻辑是硬编码式实现，不适合作为长期正式权限模型。
- 仓库目录中存在真实业务数据、上传材料、证书文件和服务器信息，这些内容不应继续按公开仓库方式直接分发。

建议至少执行以下治理动作：

- 将 `data/`、`uploads/`、证书文件从版本管理中移出。
- 将服务器地址、密码、证书路径、部署私密脚本改为环境变量或私有部署仓库存放。
- `admin_users.json` 仅保留 `password_sha256`，停止使用明文密码。
- 为上传目录、数据目录建立独立备份策略。

## 参考文件

- [api.py](/Users/doormi/Documents/time-test1/api.py)
- [config.py](/Users/doormi/Documents/time-test1/config.py)
- [services/excel_service.py](/Users/doormi/Documents/time-test1/services/excel_service.py)
- [index.html](/Users/doormi/Documents/time-test1/index.html)
- [DEPLOYMENT.md](/Users/doormi/Documents/time-test1/DEPLOYMENT.md)
