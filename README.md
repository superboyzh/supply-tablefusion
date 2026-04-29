# 表格转换工具（Go + Vue3）

本项目是一个本地运行的 Excel 转换工具。Go 后端内嵌 Vue3 前端页面，用户双击启动脚本后自动打开浏览器使用。

## 技术栈

| 模块 | 技术 |
| --- | --- |
| 前端 | Vue3 + Vite + Element Plus |
| 后端 | Go 1.26.2 |
| Excel 处理 | github.com/xuri/excelize/v2、github.com/shakinm/xlsReader |
| 打包 | Go 编译单文件，内嵌前端页面和硬件产品映射表 |
| 启动 | `start.bat` / `start.sh` 一键启动并自动打开浏览器 |

## 核心功能

- 支持选择 `出库表` / `微店表`。
- 支持单文件上传：转换完成后直接下载一个 `.xlsx`。
- 支持批量上传：多个文件转换完成后下载一个 `.zip`，zip 内只包含转换后的 `.xlsx` 文件。
- 出库表支持 `.xls` / `.xlsx`，按 `示例文件/硬件产品信息.xlsx` 做产品名称到货品名称的映射。
- 微店表支持 `.xlsx`，会过滤 `订单状态=已关闭` 和 `订单状态=待付款` 的订单，并按商品 ID 汇总配件数量；未配置到输出列的商品会按“商品名称 *数量”写入备注。
- 单文件转换会在本机 `logs/` 目录生成 Markdown 排查日志；批量 zip 内不包含日志。
- 上传文件只在本机内存中处理，不做持久化保存。

## 当前微店商品 ID 映射

| 配件 | 商品 ID |
| --- | --- |
| 墨盒 | `4722165469`、`7255807856` |
| 墨盒海绵 | `7316226980`、`4880927524` |
| 章环 | `4466273920` |
| 定位卡片 | `4478885116` |
| 铜章印油 | `4294795757` |
| 光敏印油 | `4402625197` |
| 工作台垫 | `4295433136` |
| 手动版墨盒 | `6110548102` |
| 智能章底盖 | `4294793637` |
| 光敏底座 | `4294778913` |
| 垫片 | `4295441048` |
| 3M胶 | `4295371280` |
| 环形胶 | `4339144891` |
| 定制章环 | `6121614436` |
| 定制光敏底座 | `6239519854` |
| 木工胶 | `4294767099` |
| 印章包 | `4458446275` |
| 工作台电源 | `4517103559` |
| 印章电源 | `4425522452` |
| 光敏配件一套 | `4466313670` |
| 铜章配件一套 | `4465501827` |
| 光敏章快拆配件 | `7244521801` |
| 铜章快拆配件 | `7245501786` |

## 项目结构

```text
supply-tablefusion/
├── web/                    # Vue3 前端
├── internal/
│   └── excel/              # Excel 解析和转换逻辑
├── 示例文件/                # 样表、参考输出、硬件产品映射表
├── build-output.sh         # Mac/Linux 一键构建 Windows 交付目录
├── build-output.bat        # Windows 一键构建 Windows 交付目录
├── main.go                 # Go 程序入口
├── start.bat               # Windows 双击运行
├── start.sh                # Mac/Linux 运行
├── go.mod
└── README.md
```

## 交付物

Windows 用户只需要这两个文件放在同一个目录：

```text
表格转换工具/
├─ start.bat
└─ app.exe
```

使用方式：双击 `start.bat`，程序会启动本地服务并自动打开浏览器。

推荐生成的 Windows 交付目录：

```text
output/表格转换工具/
├─ start.bat
└─ app.exe
```

## 本机运行和打包

### 前置要求

- Go 1.26.2
- Node 24，当前项目按 `fnm use 24` 使用 Node 24
- npm

### 一键构建 Windows 交付目录

如果你改了前端、后端或转换规则，想重新生成给 Windows 用户使用的交付文件，执行一键脚本即可。

Mac/Linux：

```bash
chmod +x build-output.sh
./build-output.sh
```

Windows：

```bat
build-output.bat
```

两个脚本都会自动完成：

1. 切换 Node 24（如果本机安装了 `fnm`）。
2. 安装前端依赖。
3. 构建 Vue 前端页面。
4. 整理 Go 依赖。
5. 运行 Go 测试。
6. 编译 Windows `app.exe`。
7. 输出交付目录：

```text
output/表格转换工具/
├─ start.bat
└─ app.exe
```

然后把 `output/表格转换工具/` 复制到 Windows 电脑，双击 `start.bat` 测试。

## 修改后如何重新构建

最简单方式是直接执行一键脚本，它会重新构建前后端，并刷新 `output/表格转换工具/`。

Mac/Linux：

```bash
./build-output.sh
```

Windows：

```bat
build-output.bat
```

### 常见修改场景

只改前端、后端、Excel 转换逻辑或依赖时，都执行同一个一键脚本即可。

Mac/Linux：

```bash
./build-output.sh
```

Windows 上执行：

```bat
build-output.bat
```

## 常见问题

### 下载文件名乱码

已使用标准 `filename*` UTF-8 响应头处理中文下载名。如果浏览器仍显示异常，优先使用 Chrome / Edge 测试。

### 端口被占用

默认端口是 `18080`。如果被占用，可以指定端口启动：

```bash
PORT=18081 ./start.sh
```

Windows 可临时修改 `start.bat` 里的：

```bat
set PORT=18081
```

### Windows 提示安全风险

这是本地编译的未签名 exe 常见提示。点击“更多信息”后选择“仍要运行”即可。
