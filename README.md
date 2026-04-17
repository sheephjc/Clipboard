# Clipboard 1.0.0

一个剪贴板历史小程序。当前 Windows 版已可用，macOS 版源码适配已开始准备。

## 工程结构

- `clipboard_manager.py`：共享 Tk UI 和当前 Windows 入口
- `shared\\`：跨平台数据模型、便携数据目录和富文本辅助代码
- `platforms\\windows\\`：Windows 平台适配入口，当前复用已验证的 pywin32 实现
- `platforms\\macos\\`：macOS 平台适配源码，基于 PyObjC/NSPasteboard，待 Mac 真机验证
- `assets\\`：图标资源

## 当前能力

- 监听系统剪贴板，并按 `文本 / 图片 / 其他 / 收藏` 分类保存
- 文本内容支持编辑、富文本复制与纯文本复制
- 图片内容保存到本地并可重新写回系统剪贴板
- 文件和文件夹复制会进入“其他”分类并展示路径列表
- 默认以托盘程序运行，关闭窗口时隐藏到托盘而不是退出
- 支持通过注册表实现当前用户开机自启
- 自动将旧数据目录 `%APPDATA%\\Clipboard` / `%APPDATA%\\ClipboardTrayApp` 复制迁移到程序目录旁的 `data`

## 开发运行

Windows：

```powershell
pip install -r requirements.txt
python clipboard_manager.py
```

程序启动后默认驻留托盘，左键托盘图标可显示或隐藏主窗口。

macOS 源码准备：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-macos.txt
python clipboard_manager.py
```

macOS 版需要在 Mac 上验证剪贴板、菜单栏和打包行为。

## PyInstaller 打包

Windows：

```powershell
pip install pyinstaller
pyinstaller clipboard_manager.spec
```

打包完成后会生成：

- `dist\\Clipboard 1.0.0\\Clipboard.exe`

分发时请发送整个 `dist\\Clipboard 1.0.0` 文件夹，而不是只发单个 `Clipboard.exe`。

macOS：

```bash
pyinstaller clipboard_macos.spec
```

该命令需在 macOS 上执行，生成的 `.app` 还需要在 Mac 上做真实运行验收。

## GitHub Pages 发布

发布页放在 `docs\\index.html`，可在 GitHub 仓库的 Settings → Pages 中选择 `main / docs` 启用。

下载文件不要提交进 Git 仓库。推荐流程：

1. 打包 Windows 版，确认 `dist\\Clipboard 1.0.0` 里没有 `data\\`。
2. 生成 `release\\Clipboard-Windows.zip`，压缩内容只包含 `Clipboard.exe` 和 `_internal\\`。
3. 在 GitHub Releases 创建新版本，例如 `v1.0.0`。
4. 上传 `Clipboard-Windows.zip` 作为 release asset。
5. macOS 版完成后上传同名 `Clipboard-macOS.zip`，发布页按钮会沿用同一套链接规则。

仓库里放源码和 `docs\\` 发布页；二进制 zip 放 GitHub Releases。`release\\`、`dist\\`、`build\\` 和 `data\\` 都会被 `.gitignore` 排除。

## 数据目录

- 数据库：程序目录旁的 `data\\history.db`
- 图片缓存：程序目录旁的 `data\\images\\`
- 首次运行新版时，会从旧目录 `%APPDATA%\\Clipboard` 或 `%APPDATA%\\ClipboardTrayApp` 复制迁移已有数据。
- 发布压缩包不得包含 `data\\`、`history.db` 或 `images\\`。
