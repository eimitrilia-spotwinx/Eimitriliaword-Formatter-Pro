# Eimitriliaword Formatter Pro

面向中文正式文档的**本地 Word 自动排版工具**。把排版规范写进参数，把文档拖进来，按规范重排一遍——原文件不动，结果另存。

适用于党政机关公文、企业报告、学术论文等对格式有硬性要求的场景。

> 本项目基于 [cwyalpha/Word-Formatter-Pro](https://github.com/cwyalpha/Word-Formatter-Pro)（MIT License）的排版内核，在此之上重做了 Web 界面、本地服务层与打包流程。

![界面预览](assets/sig-footer.png)

## 特性

- **完全本地**：文档不上传，无遥测，不联网校验
- **格式规范化**：字体字号、行距、缩进、页边距、标题层级、页眉页脚
- **表格处理**：字体字号、表头加粗、列宽自适应、边框线宽、单元格对齐
- **附件与落款**：附件标题格式、落款版式、日期规范化
- **中英混排**：中文用中文字体，英文数字用西文字体
- **批量处理**：多文件拖入，或整文件夹递归读取
- **三套预设**：党政机关公文（默认）、企业报告、学术论文
- **83 项参数**：常用 26 项直接可见，高级 57 项按 9 类分组抽屉式管理
- **界面可调**：左右分栏、拖拽区高度、日志高度、抽屉宽度都可拖拽并记忆
- **深色模式**：跟随系统 / 浅色 / 深色 三态切换

## 快速开始

### 方式一：直接运行源码（推荐）

```bash
git clone https://github.com/eimitrilia-spotwinx/Eimitriliaword-Formatter-Pro.git
cd Eimitriliaword-Formatter-Pro
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python wfp_web.py
```

浏览器会自动打开 `http://127.0.0.1:8878/`。

### 方式二：使用打包好的 exe

从 [Releases](https://github.com/eimitrilia-spotwinx/Eimitriliaword-Formatter-Pro/releases) 下载 zip，解压后双击 `Eimitriliaword Formatter Pro.exe`。

> 未签名，Windows 可能弹出 SmartScreen 提示，选「更多信息 → 仍要运行」。

## 环境要求

| 项 | 要求 |
|---|---|
| 系统 | Windows 10 / 11 |
| Python | 3.11 及以上（开发用 3.14） |
| 依赖 | `python-docx`、`lxml`、`fonttools`、`pywin32` |
| Office | 处理 `.doc` / `.wps` 需要本机安装 **WPS** 或 **Microsoft Word**；`.docx` 不需要 |

## 支持格式

`.docx` · `.doc` · `.wps` · `.txt` · `.md`

## 使用流程

1. 把文件（或文件夹）拖进左侧虚线框
2. 右侧选预设，或逐项调整参数；高级项在右上角「高级设置」
3. 确认底部「输出到」目录，点「开始排版」（`Ctrl+Enter`）
4. 完成后点「打开目录」取文件

## 构建 exe

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

脚本会自动完成：停止旧实例 → 准备 3.11 打包环境 → 依赖自检 → 生成版本信息 → PyInstaller onedir 构建 → 同步前端。

产物：`dist\Eimitriliaword Formatter Pro\Eimitriliaword Formatter Pro.exe`（约 39 MB / 129 文件，首次约 27 秒，之后约 12 秒）。

**只改前端不必重打包**：把 `webui\index.html` 复制到 `dist\...\_internal\webui\`，窗口按 F5 即可。

## 技术栈

| 层 | 实现 |
|---|---|
| 前端 | 单文件 HTML + 原生 JS，零外部依赖，无框架 |
| 后端 | Python 标准库 `http.server`（`ThreadingHTTPServer`） |
| 文档处理 | `python-docx` + `lxml` |
| 格式转换 | 通过 COM 调用本机 WPS / Word |
| 界面外壳 | Edge `--app` 无地址栏窗口 |
| 打包 | PyInstaller onedir |

## 项目结构

```
├─ wfp_web.py                 后端主程序（HTTP 服务、接口、窗口拉起）
├─ webui/index.html           前端全部界面（单文件）
├─ wfp_core.py                排版内核
├─ wfp_config.py              默认配置定义
├─ wfp_cli.py                 命令行入口
├─ build.ps1                  一键打包脚本
├─ packaging/
│   └─ make_version_info.py   生成 exe 版本资源
├─ assets/                    图标与截图
└─ 启动.bat                    兜底启动入口
```

## 已知限制

- `.doc` / `.wps` 依赖本机 Office，未安装则无法处理
- exe 未签名，可能被 SmartScreen 拦截
- 字体缺失时 Word 会自动替换，界面会提示但不阻止
- COM 转换串行执行，大文档需等待
- 界面日志最多保留 400 行
- 点「取消」会跳过后续文件，但当前文件会跑完
- 同目录下同名文件重复排版会覆盖

## 版本

当前版本：`2.7.9-eimitrilia.1`

命名规则：`<上游内核版本>-eimitrilia.<定制序号>`。版本号唯一真源在 `wfp_version.py`。

## 许可证

MIT License。原创内核版权归 [cwyalpha](https://github.com/cwyalpha) 所有，本项目在此基础上修改并再发布。

详见 [LICENSE](LICENSE)。

## 作者

**Eimitrilia Spotwinx**

© 2026 Eimitrilia Spotwinx
