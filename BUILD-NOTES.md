# Eimitriliaword Formatter Pro · Web 界面构建笔记

> **当前交付形态（2026-11-25 晚 定版）：三条路，按推荐度排序**
>
> 1. **桌面快捷方式（推荐，在用）**：`桌面\Eimitriliaword Formatter Pro.lnk` → `pythonw.exe wfp_web.py`。
>    没有 exe，因此**完全不触发 SmartScreen**；pythonw 无控制台窗口。
> 2. **exe（onedir）**：`dist\Eimitriliaword Formatter Pro\Eimitriliaword Formatter Pro.exe`，双击开窗、关窗即退。
>    **但会被 SmartScreen 拦**（见第八节），仅作备用。
> 3. **源码兜底**：双击 `启动.bat` → 本地服务 + 浏览器。
>
> **已放弃 onefile**：`--runtime-tmpdir` 不展开环境变量，只能用绝对路径，exe 绑死本机；且每次启动解压约 200 MB。
> **Edge 空白窗必须加 `--no-sandbox`**：由 Python 进程拉起时 Chromium 渲染沙箱初始化失败，详见第六节。
> **自签名证书对 SmartScreen 无效**，已撤销，详见第八节。
> 改 HTML 不用重打包：直接改 `dist\Eimitriliaword Formatter Pro\_internal\webui\index.html` 后刷新窗口即可（快捷方式形态则改源码里的 `webui\index.html`）。

> 记录于 2026-11-25。目标产物：**onedir exe，双击开窗，关窗即退**（已从单文件改为目录形态）。
> 本文档的目的：下次改完界面重打包时，不用再把坑重新踩一遍。

---

## 一、成品与源码位置

| 项 | 路径 |
|---|---|
| 成品（onedir） | `dist\Eimitriliaword Formatter Pro\Eimitriliaword Formatter Pro.exe`（5.8 MB，整包 39.1 MB / 129 文件，未签名） |
| 后端 | `wfp_web.py`（标准库 `http.server`，含窗口启动逻辑） |
| 前端 | `webui\index.html`（单文件，零外部依赖） |
| 源码环境（日常运行） | `_env\.venv`（Python 3.14，桌面快捷方式跑这个） |
| 打包环境 | `_build\.venv`（Python 3.11，`--system-site-packages`，只补装 pywin32） |
| 构建脚本 | `build.ps1` |
| 版本信息生成 | `packaging\make_version_info.py` |
| 上游源码 | 本目录即 clone 的仓库根，**上游文件未做任何修改** |

运行期数据（打包版）：
- 输出：`%LOCALAPPDATA%\WordFormatter\out`（不可写时回退 exe 同目录 → 系统临时）
- 日志：`%LOCALAPPDATA%\WordFormatter\app.log`
- 临时：`...\WordFormatter\tmp`（同时被设为 TMP/TEMP 与 gen_py 根）

---

## 二、构建命令（可直接复制）

```powershell
cd <项目目录>
powershell -ExecutionPolicy Bypass -File .\build.ps1

# 可选
.\build.ps1 -KeepData     # 保留 dist 下的运行时数据目录
```

`build.ps1` 自动完成六件事：

1. 停掉运行中的实例（否则 dist 被占用）
2. 复用/创建 3.11 打包环境（首次约 27 秒，之后 12 秒）
3. 依赖自检：docx / lxml / fontTools / win32com / PyInstaller
4. 生成 `version_info.txt`（版本与作者取自 `wfp_version.py`）
5. PyInstaller onedir 构建，带图标与版本资源
6. 同步前端到 `_internal\webui\` 并校验哈希

**只改前端时不必重打包**：把 `webui\index.html` 复制到
`dist\Eimitriliaword Formatter Pro\_internal\webui\`，窗口按 F5 即可。

---

## 三、环境事实（本机特有，务必记住）

1. **Python 子进程写不了 `C:\Users\<用户名>\AppData\Local\Temp`**（`WinError 5`），
   但 PowerShell / .NET 能写。ACL 正常（Owner = 当前用户，FullControl），属进程级限制。
   → 所有临时文件、`pywin32` 的 `gen_py`、onefile 解压都受影响。
2. **DSH 的子进程写工作区外会被拒**，即使文件策略是 `danger-full-access`。
   → **在我这边跑通 ≠ 用户双击能跑通**，结论必须标明测试环境。
3. **PyPI 直连不可靠**（15 s 超时、多次重试）；清华源 15 秒装完全部依赖。
4. **npm 的 `EPERM` 曾被我误判为"网络坏了"**：实际是写缓存被拒，
   `EPERM` 被包在 `FetchError` 里，外层文案极具误导性。

---

## 四、踩过的坑（21 条）

### A. 环境层

| # | 现象 | 真相 / 对策 |
|---|---|---|
| 1 | `WinError 5` 写 `%TEMP%` | 进程级限制 → 重定向 TMP/TEMP，并**用探针文件试写**，不能只 `mkdir` |
| 2 | 子进程写工作区外被拒 | 沙箱行为 → 测试结论必须区分「沙箱内」与「用户双击」 |
| 3 | PyPI 超时 | 构建期依赖直接走清华源（包管理器安装仍守默认源规则） |
| 4 | `npm EPERM` 误判 | 先看 `syscall` + 路径，再看外层错误文案 |

### B. PyInstaller 层

| # | 坑 | 对策 |
|---|---|---|
| 5 | onefile 解压到系统 TEMP 失败：`Could not create temporary directory!` | `--runtime-tmpdir` 指到固定可写目录 |
| 6 | **`pyi_rth_win32comgenpy` 运行时钩子**在业务代码之前就把 `win32com.__gen_path__` 设成系统 TEMP | 导入 `wfp_core` 前显式覆盖 `win32com.__gen_path__`（本项目已实现） |
| 7 | `--noconsole` 下 `sys.stdout is None`，`print()` 抛异常 | 统一走 `say()` → 日志文件 |
| 8 | `--noconsole` 下 bootloader 错误不可见 | **先打 `--console` 版定位问题，再切 noconsole** |
| 9 | frozen 后路径语义混乱 | 只读资源 = `sys._MEIPASS`；可写数据 = exe 目录 / `%LOCALAPPDATA%` |
| 10 | exe 可能被放在只读目录 | 可写目录三级回退：LocalAppData → exe 同目录 → 系统临时 |

### C. pywebview / pythonnet 层（结论：别用）

| # | 坑 | 说明 |
|---|---|---|
| 11 | `Failed to resolve Python.Runtime.Loader.Initialize` | 默认 netfx 加载失败；设 `PYTHONNET_RUNTIME=coreclr` + `PYTHONNET_DOTNET_ROOT` 可解 |
| 12 | pywebview 新版依赖 `proxy-tools`（连字符），清华源没有 | 镜像不全时安装必卡 |
| 13 | `--only-binary=:all:` 触发 pip 版本回溯 | 800+ CPU 秒空转；改用 `--no-deps` + 手动补依赖 |
| 14 | 版本号要精确：`clr_loader==0.2.7` 不存在，是 `0.2.7.post0` | — |
| 15 | **最终选择**：Edge `--app` 模式 | 零依赖、无地址栏、关窗即退，效果与 pywebview 接近 |

窗口启动参数（当前实现）：

```
msedge.exe --app=http://127.0.0.1:8878/ --user-data-dir=<data>\window-profile
           --window-size=1200,860 --no-first-run --no-default-browser-check
```

`Popen(...).wait()` 阻塞到窗口关闭 → 再 `httpd.shutdown()`，实现「关窗即退」。

### D. 与上游代码集成

| # | 坑 | 对策 |
|---|---|---|
| 16 | `normalize_format_options()` 只做迁移校验，**不补默认值** → `KeyError: 'set_outline'` | 调用方必须 `merged = dict(DEFAULT_CONFIG); merged.update(user_cfg)` |
| 17 | COM 必须**每线程**初始化，而 HTTP 每请求一线程 | 处理函数内 `_initialize_com_for_thread()`，`finally` 里释放 |
| 18 | `.doc/.wps` 需 pywin32 + 本机 WPS/Word；`.docx` 不需要 | 文档中分开说明 |

### E. 流程层

| # | 坑 | 对策 |
|---|---|---|
| 19 | "打开就能用"被理解成"能跑起来" | 需求先拆硬指标：单文件 / 弹窗 / 无控制台 / 关窗退出 |
| 20 | 验收覆盖不全 | 固定五项，见下 |
| 21 | 全局规则里 pip 被排除镜像 | 实测 PyPI 直连不可用，建议（待批准）放宽 |

---

## 五、发布前验收清单（五项全过才算完成）

```powershell
$exe = "dist\Eimitriliaword Formatter Pro\Eimitriliaword Formatter Pro.exe"

# 1 启动：进程存活 + 端口监听
# 2 元数据
curl.exe -sS http://127.0.0.1:8878/api/meta
# 3 .docx 排版
# 4 .doc 排版（验证 pywin32 确实打进包里）
# 5 关窗后 exe 进程退出
```

实测基线（2026-11-25）：

| 项 | 结果 |
|---|---|
| 启动 | < 1 s |
| `/api/meta` | 200，v2.7.9-eimitrilia.1 |
| `.docx` | 38,764 字节 |
| `.doc` | 25,897 字节 |
| 关窗退出 | 4 s 内进程消失 |

---

## 六、运行时窗口问题（2026-11-25 晚追加，两个都踩过）

### 6.1 `--noconsole` 下写访问日志会断连接

**现象**：`/` 返回 200，`/api/meta` 直接 `Empty reply from server`（curl 52），
前端拿不到配置，界面停在"连接中"。

**原因**：`--noconsole` 打包后 `sys.stderr` 为 `None`；`BaseHTTPRequestHandler.log_message`
里 `sys.stderr.write(...)` 抛 `AttributeError`，此时响应头还没发出 → 连接被掐断。
当时的代码只对 `/api/*` 写日志，所以只有 API 挂、首页正常——极具迷惑性。

**修法**：`log_message` 整体静默（或先判 `sys.stderr` 是否为 None），桌面端不需要访问日志。

### 6.2 由 Python 启动的 Edge，窗口出现但整片空白

**现象**：窗口标题停在 `127.0.0.1_/`（URL 而非页面标题），内容全灰，服务端一切正常。

**排查过程**（这段最重要，下次照做可省 40 分钟）：
1. 服务端自检：`curl /` 与 `curl /api/meta` 都 200 → 排除服务。
2. 用 Edge 无头模式 `--dump-dom` 取同一 URL → DOM 完整 → 排除"Edge 连不上"。
3. A/B：用**PowerShell 的 `Start-Process`** 启动同参数 Edge → **正常**；
   用 **Python 的 `Popen` / `os.startfile` / `cmd start`** 启动 → **全部空白**。
   → 关键变量是"谁去启动 Edge"，不是参数、不是路径、不是环境变量。
4. 二分 Chromium 开关：`--disable-gpu`、`--disable-gpu-sandbox`、`--in-process-gpu`、
   `--disable-features=RendererCodeIntegrity` **全无效**；**只有 `--no-sandbox` 恢复渲染**。

**根因**：被 Python（尤其是冻结后的 exe）拉起时，Chromium 的渲染进程沙箱初始化失败，
渲染器起不来，于是窗口在、页面空。

**修法**：启动参数加 `--no-sandbox`。
本项目页面只加载 `127.0.0.1` 上的自有资源、无任何外部内容，风险可控；
若将来要加载外部页面，需重新评估这条。

**验证口径**：窗口标题 == 页面标题（"Eimitriliaword Formatter Pro"）即代表 HTML 已解析并渲染；
停在 URL 标题即为失败。

---

## 七、已知限制

- **未签名**：首次运行触发 SmartScreen。
- **窗口标题**：Edge 刚打开时可能短暂显示 `127.0.0.1_/`，页面加载后应变为应用名。
- **依赖本机 Edge**：找不到时退回默认浏览器（并写日志）。
- **处理串行**：全局锁，防 COM 与临时文件竞争；并发请求排队。

---

## 八、SmartScreen 与签名（2026-11-25 实测结论）

**结论：自签名证书无法通过 SmartScreen。**

实测过程：

1. 生成自签名代码签名证书（Code Signing EKU、SHA256、私钥不可导出）
2. 导入「受信任的根证书颁发机构」+「受信任的发布者」（CurrentUser）
3. `Set-AuthenticodeSignature` 签名，`Get-AuthenticodeSignature` 返回 **Valid**，带时间戳
4. 双击运行 → **仍然弹出**「Windows 已保护你的电脑 / SmartScreen 阻止了无法识别的应用启动」
5. 点「更多信息 → 仍要运行」后，**下次双击依旧弹**

已排除的其他原因：

| 检查项 | 结果 |
|---|---|
| MOTW（Zone.Identifier） | 不存在 |
| 智能应用控制 Smart App Control | 已关闭（VerifiedAndReputablePolicyState=0） |
| WDAC / 应用控制策略 | 0 条 |
| Defender 实时保护 | 开启（正常，不是拦截源） |

**原因**：SmartScreen 查的是**微软云端的文件声誉**，不是本机证书链。
自签名证书不在微软信任列表里，对云端等于零声誉。
本地信任根只影响「未知发布者」那类提示，管不到 SmartScreen。

三条可行路径：

| 方案 | 效果 | 代价 |
|---|---|---|
| **快捷方式 → pythonw.exe** | 完全不触发（pythonw 是官方签名、有声誉） | 图标不是 .exe |
| 关闭「检查应用和文件」 | 不再弹 | 全局降低防护 |
| EV 证书 / Azure Trusted Signing | 正规解 | 花钱 + 身份验证 |

**本项目采用**：桌面快捷方式 `Eimitriliaword Formatter Pro.lnk` →
`_env\.venv\Scripts\pythonw.exe wfp_web.py`。

证书已从三个存储（My / Root / TrustedPublisher）撤销，`.cer` 已删除；
`dist\Eimitriliaword Formatter Pro\Eimitriliaword Formatter Pro.exe` 的签名因此变成「链终止于不受信任的根」，
仅在确实需要 exe 形态时使用。
---

## 九、版本号与发布信息（单一来源）

**唯一真源**：`wfp_version.py` 里的 `__version__`。

命名规则：`<上游内核版本>-eimitrilia.<定制序号>`

例：`2.7.9-eimitrilia.1` = 基于上游 wfp 2.7.9 的第 1 版定制。

### 升版本只做两件事

1. 改 `wfp_version.py`：

   ```python
   __version__ = "2.7.9-eimitrilia.2"     # 只动后缀数字
   ```

2. 重新打包：

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build.ps1
   ```

### 自动同步到哪些地方

| 位置 | 机制 |
|---|---|
| 页面 / 窗口标题栏 | `<span data-ver>` 由 `/api/meta` 的 `version` 填充 |
| 关于页 4 处（抽屉栏 / 副标题 / 当前版本 / 页脚） | 同上 |
| 启动日志 | `wfp_web.py` 打印 `APP_NAME v__version__` |
| exe 文件属性 | `packaging\make_version_info.py` → `version_info.txt` → `--version-file` |

**禁止**在前端 HTML 里硬编码版本号。需要显示时统一用：

```html
<span data-ver></span>
```

### exe 文件属性字段

| 字段 | 值 |
|---|---|
| ProductName / InternalName | `Eimitriliaword Formatter Pro` |
| CompanyName | `Eimitrilia Spotwinx` |
| LegalCopyright | `© 2026 Eimitrilia Spotwinx` |
| FileDescription | `Word 文档智能排版工具` |
| FileVersion / ProductVersion | `__version__` 原样；数字段自动换算为 `(a, b, c, d)` |

例：`2.7.9-eimitrilia.1` → 数字段 `(2, 7, 9, 1)`。
