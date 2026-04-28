# 表格转换工具（Go + Vue3）

## 技术栈

| 模块 | 技术 |
| --- | --- |
| 前端 | Vue3 + Vite + Element Plus |
| 后端 | Go 1.26.2（单文件 exe，无依赖） |
| Excel 处理 | github.com/xuri/excelize |
| 打包 | Go 直接编译成 exe（内置前端页面） |
| 启动 | bat / sh 一键启动 + 自动开浏览器 |

## 核心功能

- 用户选择：出库表 / 微店表
- 上传 Excel
- Go 后端自动字段转换
- 下载处理后的标准表格
- 全程不存数据、用完即走

## 交付物

> 用户电脑零安装，双击 `.bat` / `.sh` 即用。

```text
表格转换工具/
├─ start.bat       # Windows 双击运行
├─ start.sh        # Mac/Linux 运行
└─ app.exe         # 前后端一体，无任何依赖
```

**用户操作流程：** 双击 `start.bat` → 自动打开浏览器 → 直接用

## 项目结构

```text
supply-tablefusion/
├── web/              # Vue3 前端
├── internal/         # Go 后端逻辑
│   └── excel/        # Excel 处理
├── main.go           # Go 程序入口
├── go.mod
└── 打包脚本
```
