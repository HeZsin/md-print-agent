# 明道云静默打印代理（md-print-agent）

让明道云（HAP）记录一键打印到本地打印机：全程无弹窗、无对话框、无需用户手动操作。
单文件 Python 脚本（零 pip 依赖，仅标准库），可打包为单 exe 分发给客户机。

## 架构原理

核心问题：明道云在云端，打印机在本地（NAT 后）——云端永远无法主动触达本地电脑。
解法：方向反转，本地代理**主动轮询拉取**（出站请求永远通）。

```
明道云业务表 ──自定义按钮──▶ 工作流：获取记录打印文件 → 新建记录到「待打印队列」
                                        │
本地代理 watch 模式：每 3s 轮询队列表 ──┘
  捞「待打印」→ 置「打印中」→ 下载附件 → 转 PDF → SumatraPDF 静默出纸 → 回写「已打印/失败」
```

任务统一汇入一张队列表的设计：业务表零侵入（不改状态、不加字段），任意业务表复用同一套打印链路，本地代理只需盯一张表。

## 两种工作模式

| 模式 | 触发方式 | 适用场景 |
|------|---------|---------|
| **watch（推荐）** | 业务表按钮 → 工作流渲染 Word/Excel 模板 → 队列表 | 打印模板固定、按记录出单 |
| HTTP | 油猴脚本「明道云静默打印助手」hook 页面打印内容 POST 到本代理 | 打印页面级 HTML（浏览器内所见） |

私有部署（同内网）可选实时推送：明道云 Webhook 直打本代理 `POST /print`。

## 明道云侧配置（标准工作流）

业务表任意，任务统一汇入「待打印队列」表：

1. **自定义动作按钮**触发
2. 节点1「获取记录打印文件」：打印对象 = 按钮触发的工作表；打印模板/文件名按业务配置
   （Word/Excel 模板免费，PDF 模板收费 0.15 元/份）
3. 节点2「新建记录」：选择工作表 = 待打印队列，新增一条记录
   - 名称 = 自定义；打印编号 = 留空（自动编号）
   - 打印状态 = 待打印；附件 = 节点1 输出的打印文件

「待打印队列」表字段规范：

| 字段 | 类型 | 说明 |
|------|------|------|
| 名称 | 标题/文本 | 自定义记录名称 |
| 打印编号 | 自动编号 | 规则 print-自然数（前缀需界面设置，API 建不了） |
| 打印状态 | 单选 | 待打印 / 打印中 / 已打印 / 失败 |
| 附件 | 附件字段 | 工作流写入的打印文件 |

代理首次运行向导可选自动建表（含上述字段，别名 num/status/file 自动对齐）。

## 客户机部署

```
md-print-agent.exe      # 打包见下方
SumatraPDF.exe          # Portable 版，与 exe 同目录（或装到默认路径）
config.json             # 首次运行向导自动生成
```

首次运行 `md-print-agent.exe --watch` 进入交互式向导：明道云 API 地址/AppKey/Sign →
建表或填已有工作表 → 选打印机 → 轮询间隔，自动校验连接并保存配置，随后直接进入轮询。

本地依赖（自动探测）：
- **PDF 打印**：SumatraPDF（https://www.sumatrapdfreader.org）
- **Word/Excel → PDF**：MS Office 或 WPS（COM 调用，双引擎兜底）
- **HTML → PDF**（仅 HTTP 模式）：系统自带 Edge/Chrome 无头渲染

## 用法（源码方式）

```
python print_agent.py --watch          # watch 模式（推荐）；首次运行自动进配置向导
python print_agent.py --test           # 自检：打一张测试页到默认打印机
python print_agent.py                  # HTTP 模式（默认 127.0.0.1:8765，配油猴脚本）
python print_agent.py --watch --config D:\mdprint.json
```

打包单 exe：

```
python -m PyInstaller --onefile --name md-print-agent print_agent.py
```

## 调试通路的方法

**不要为客户创建工作流**——直接在「待打印队列」表 API 创建一条测试记录即可：
状态 = 待打印 + 附件 = 测试 PDF（base64），watch 捞到后走完整管线，从日志和状态回写判断链路是否通。这样验证的是代理侧全链路，工作流只是往队列表写数据的生产者，与链路无关。

## 踩坑经验（V3 API 实测，与官方文档不符处）

1. **更新记录必须用 PATCH**：`PATCH /v3/app/worksheets/{id}/rows/{row_id}`，POST/PUT 均返回 405（文档写 POST 是错的）。
2. **单选字段直接传选项文本值**：筛选和更新都无需选项 key，服务端自动匹配（文档要求传 key，实测不必要）。
3. **建表 `POST /v3/app/worksheets` 无需 appId**：AppKey 鉴权自带应用归属。
4. **AutoNumber 编号规则 API 传不了**：只能建纯数字递增，`print-` 前缀规则需在界面设置。
5. **回写状态必须 `triggerWorkflow: false`**：否则渲染→打印→回写会触发工作流死循环。
6. **附件异步**：写入后 5-10 秒才可下载；下载用返回的 `downloadUrl`（预签名地址，免鉴权）。
7. **附件字段名兼容**：返回结构用驼峰 `fileName`/`downloadUrl`，不同版本字段名不一致，需做兜底链。
8. **PyInstaller onefile 路径坑**：exe 模式下 `__file__` 指向解压临时目录，必须 `sys.frozen` 判断后取 `sys.executable` 所在目录为基准（config.json / SumatraPDF 放 exe 旁才找得到）。
9. **PowerShell 中文输出乱码**：枚举打印机等场景需强制 `[Console]::OutputEncoding=UTF8` 并按字节捕获显式解码，不依赖系统区域设置。
10. **Edge 无头打印拿不到文件**：Edge 已在运行时无头进程会把任务转交常驻实例，必须用独立 `--user-data-dir`。
11. **打印机枚举含脱机副本**：同名打印机可能有多个条目（如"xxx (副本 1)"），config 的 `printer` 必须填准确名称，留空走系统默认需确认默认打印机在线。

## 文件说明

| 文件 | 说明 |
|------|------|
| `print_agent.py` | 代理主程序（HTTP + watch 双模式，含配置向导与自动建表） |
| `config.example.json` | 配置模板（真实 config.json 不入库） |
| `明道云静默打印助手.user.js` | 油猴脚本（HTTP 模式的浏览器端） |
