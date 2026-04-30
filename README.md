# Clipboard 2.0.2

Clipboard 是一个轻量的剪贴板历史工具，用来找回、编辑和重新复制曾经复制过的内容。它默认在后台驻留托盘，适合写作、整理资料、处理图片和反复复制文本时使用。

## 功能亮点

- 记录文本、富文本、图片、文件和文件夹路径。
- 支持按 `全部 / 文本 / 图片 / 其他 / 收藏` 分类查看历史。
- 文本记录可以编辑后重新复制，支持富文本复制和纯文本复制。
- 图片记录可以预览、重新写回剪贴板，并支持 OCR 文字识别。
- 支持搜索、收藏、日期筛选、自动清理和开机自启。
- 关闭窗口后仍可驻留托盘，需要时再从托盘打开。
- 历史数据默认保存在程序目录旁的 `data/` 文件夹，不上传到云端。

## 下载与使用

Windows 版可以在 GitHub Releases 下载压缩包：

```text
Clipboard-Win-2.0.2.zip
```

下载后解压到常用位置，运行 `Clipboard.exe` 即可。首次运行会在程序目录旁生成 `data/` 文件夹，用来保存历史数据库和图片缓存。更新版本时，如果想保留历史记录，把旧版本旁边的 `data/` 文件夹放到新版同一目录旁即可。

macOS 版源码已做平台适配，但仍需要在真实 Mac 上验收剪贴板、菜单栏、开机自启和打包行为。

## 项目架构

- `clipboard_manager.py`：主程序入口，包含 Tk 界面、托盘交互和应用层流程。
- `shared/`：跨平台核心逻辑，包括数据模型、SQLite 存储、数据目录、富文本处理和 OCR 封装。
- `platforms/windows/`：Windows 剪贴板、托盘和开机自启等平台服务。
- `platforms/macos/`：macOS 平台适配代码，基于 NSPasteboard、LaunchAgent 和菜单栏能力。
- `assets/`：应用图标资源。
- `docs/`：GitHub Pages 发布页和示例截图资源。
- `tests/`：核心存储和平台服务相关测试。

整体上，UI 和应用流程在主入口里组织，跨平台可复用能力放在 `shared/`，系统剪贴板和自启等差异化能力放到 `platforms/` 下隔离。

## 源码运行

Windows：

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python clipboard_manager.py
```

macOS：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-macos.txt
python clipboard_manager.py
```

## 打包

Windows 使用 PyInstaller：

```powershell
.\.venv\Scripts\pyinstaller.exe --clean clipboard_manager.spec
```


## 数据与隐私

- 数据库：`data/history.db`
- 图片缓存：`data/images/`
- `data/` 位于程序目录旁，默认不会写入 C 盘固定位置。
- `data/`、`dist/`、`release/`、虚拟环境和构建缓存都不会提交到仓库。
