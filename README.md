# 通讯录 Address Book

软件工程（FZU-MU）极限编程作业：一个支持**收藏联系人**、**一个联系人多种联系方式**、**Excel 导入导出**的通讯录 Web 应用。

前端原生 HTML/CSS/JS 无构建步骤，后端 Flask + SQLite，单人即可一条命令跑起来。

---

## 一、功能一览

| 需求 | 实现情况 | 说明 |
| --- | --- | --- |
| 收藏联系人（25 分） | ✅ | 列表与详情页均可一键星标；收藏的联系人固定排在列表最前；侧边栏「收藏」视图单独筛选 |
| 多种联系方式（10 分） | ✅ | 一个联系人可挂手机 / 邮箱 / 微信 / QQ / 地址 / 其他，数量不限（上限 12），每条可带标签（如「办公室」） |
| Excel 导出（25 分） | ✅ | 一人一行、一条信息一列；带表头样式、冻结首行、自动筛选，并附「说明」工作表 |
| Excel 导入（25 分） | ✅ | 兼容本应用导出格式与「手机/邮箱/微信 各占一列」的自定义格式；空行跳过、错误行定位提示、同名可合并 |
| 页面样式（10 分） | ✅ | 三栏布局、SVG 图标、响应式、加载/空态/Toast 反馈 |
| 云端部署（15 分，选做） | ✅ | 已部署上线：http://47.95.145.177/ |

额外做的：搜索（姓名 / 电话 / 分组 / 备注 / 标签）、分组筛选、一键复制联系方式、下载导入模板、删除二次确认、快捷键 `Ctrl+K`。

## 二、技术栈

- 后端：Python 3.10+ / Flask 3 / SQLite（标准库 `sqlite3`）
- Excel：openpyxl 3
- 前端：原生 JavaScript（ES6）、CSS 变量与 Grid/Flex 布局，无框架、无打包
- 测试：`tests/test_smoke.py`（urllib + openpyxl，28 项断言，覆盖 CRUD / 收藏 / 导入导出往返）

## 三、目录结构

```
fzu-address-book/
├── app.py                  # Flask 入口：路由与参数校验
├── db.py                   # 数据层：建表、种子数据、联系人 / 联系方式 CRUD
├── excel_io.py             # Excel 导入导出：导出样式、模板、两种表头解析
├── requirements.txt
├── start.bat / start.sh    # 一键安装依赖并启动
├── templates/
│   └── index.html          # 单页应用骨架（含 SVG 图标 sprite）
├── static/
│   ├── css/style.css       # 全部样式
│   ├── js/app.js           # 全部交互逻辑
│   └── img/logo.svg
├── tests/
│   └── test_smoke.py       # 端到端冒烟测试
├── docs/
│   ├── blog.md             # 作业博客（仓库版，图片用相对路径）
│   ├── blog-csdn.md        # 博客发布版（@[toc] + 图片上传位，直接粘 CSDN）
│   ├── MODULE_DESIGN.md    # 模块划分
│   ├── CODE_STANDARD.md    # 代码规范
│   └── images/             # 博客截图
└── data/                   # 运行期自动生成的 SQLite 数据库（已 gitignore）
```

## 四、快速开始

```bash
# 1. 装依赖（建议虚拟环境）
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

pip install -r requirements.txt

# 2. 启动
python app.py
# 浏览器打开 http://127.0.0.1:5000
```

Windows 直接双击 `start.bat`，macOS / Linux 执行 `bash start.sh` 也可以。

首次启动会自动建表并写入 5 条演示数据；想从空通讯录开始：

```bash
ADDRESS_BOOK_SEED=0 python app.py     # macOS / Linux
set ADDRESS_BOOK_SEED=0 && python app.py   # Windows CMD
```

指定端口 / 数据库位置：

```bash
PORT=8000 ADDRESS_BOOK_DB=data/demo.db python app.py
```

## 五、Excel 格式

### 导出的列

| 姓名 | 分组 | 收藏 | 公司 | 联系方式1类型 | 联系方式1内容 | … | 联系方式6类型 | 联系方式6内容 | 备注 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 奶龙 | 极限编程小队 | 是 | 青青草原大学 信息工程系 | 手机/手机 | 138 0591 2366 | … | 微信 | nailong_2024 | 极限编程结对伙伴 |

- 类型列写中文类型名（手机 / 邮箱 / 微信 / QQ / 地址 / 其他），斜杠后面是该条联系方式的标签；
- 点「下载模板」拿到的表格带两行示例，照着填就能导入。

### 导入支持的两种表头

1. **槽位格式**（本应用导出的格式）：`联系方式1类型` + `联系方式1内容` 成对出现，最多 6 组；
2. **按类型分列格式**（自己手写的表）：
   `姓名 | 分组 | 收藏 | 公司 | 手机号 | 邮箱 | 微信 | QQ | 地址 | 备注`，
   一格内有多个值用换行或 `；` `、` 分隔，会被自动拆成多条联系方式。

导入规则：

- 姓名为空的行跳过，并在结果里提示行号；
- 文件内重复的联系人（同名 + 首个联系方式相同）只导入一次；
- 同名合并开关打开时，已存在的同名联系人会被追加新的联系方式，而不是产生重复记录；
- 单次上限 2000 行，文件上限 10MB，只接受 `.xlsx`。

## 六、接口一览

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/api/contacts?keyword=&favorite=1&group=` | 列表 / 搜索 / 收藏筛选 / 分组筛选 |
| POST | `/api/contacts` | 新建（body 里 `methods` 为数组） |
| GET | `/api/contacts/<id>` | 详情 |
| PUT | `/api/contacts/<id>` | 全量更新 |
| DELETE | `/api/contacts/<id>` | 删除（级联删联系方式） |
| PATCH | `/api/contacts/<id>/favorite` | 设置 / 取消收藏 |
| GET | `/api/groups` | 分组及人数 |
| GET | `/api/stats` | 总数 / 收藏数 / 联系方式数 / 分组数 |
| GET | `/api/export` | 导出 xlsx |
| GET | `/api/template` | 下载导入模板 |
| POST | `/api/import` | 导入 xlsx（multipart：`file`、`merge`） |

## 七、测试

```bash
# 另开一个终端启动服务后：
BASE_URL=http://127.0.0.1:5000 python tests/test_smoke.py
```

覆盖：建表种子数据、新建 / 编辑 / 删除、收藏筛选、按联系方式搜索、导出列结构与行数、导出文件回导不产生重复、自定义表头导入、多值单元格拆分、错误行提示、模板下载。

## 八、部署到云服务器（选做 15 分，已上线）

**线上地址：<http://47.95.145.177/>**（阿里云轻量应用服务器 / Ubuntu 22.04 / 2vCPU 1G，gunicorn + Nginx + systemd）

完整步骤（选镜像、放行端口、上传、验证、常见问题）见 **[deploy/README.md](deploy/README.md)**，Ubuntu 一键脚本：

```bash
cd /opt/fzu-address-book
bash deploy/aliyun-ubuntu-setup.sh     # 加 swap + 装依赖 + systemd 服务 + Nginx 反代
```

脚本会自动：内存小于 2G 时建 2G swapfile、按内存决定 gunicorn worker 数、初始化数据库、注册开机自启、配置好 80 反代并做本机自检。

手工部署的核心三行：

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt gunicorn

# 生产环境不要用 flask run，用 gunicorn 常驻
gunicorn -w 2 -b 0.0.0.0:8000 app:app
```

Nginx 反代配置见 [deploy/nginx-addressbook.conf](deploy/nginx-addressbook.conf)
（注意 `client_max_body_size` 要放大到 12m，否则导入 Excel 会 413）。

Windows 服务器可以用 `waitress-serve --port=8000 app:app`。

## 九、文档

- 博客（仓库版，图片相对路径）：[docs/blog.md](docs/blog.md)
- 博客（CSDN 发布版，直接粘贴）：[docs/blog-csdn.md](docs/blog-csdn.md)
- **提交前必做清单**：[docs/提交前必做清单.md](docs/提交前必做清单.md)（仓库链接、截图、视频、CSDN 发布）
- 模块划分：[docs/MODULE_DESIGN.md](docs/MODULE_DESIGN.md)
- 代码规范：[docs/CODE_STANDARD.md](docs/CODE_STANDARD.md)
- 博客截图：[docs/images/](docs/images/)（`scripts/screenshot.py` 可重新生成；`git-log.png` 为提交历史截图）
