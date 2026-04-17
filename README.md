# Clipboard 1.0.0

一个面向 Windows 10/11 的剪贴板历史小程序，基于 `tkinter + pywin32 + Pillow`。

## 当前能力

- 监听系统剪贴板，并按 `文本 / 图片 / 其他 / 收藏` 分类保存
- 文本内容支持编辑、富文本复制与纯文本复制
- 图片内容保存到本地并可重新写回系统剪贴板
- 文件和文件夹复制会进入“其他”分类并展示路径列表
- 默认以托盘程序运行，关闭窗口时隐藏到托盘而不是退出
- 支持通过注册表实现当前用户开机自启
- 自动将旧数据目录 `%APPDATA%\\ClipboardTrayApp` 迁移到 `%APPDATA%\\Clipboard`

## 开发运行

```powershell
pip install -r requirements.txt
python clipboard_manager.py
```

程序启动后默认驻留托盘，左键托盘图标可显示或隐藏主窗口。

## PyInstaller 打包

```powershell
pip install pyinstaller
pyinstaller clipboard_manager.spec
```

打包完成后会生成：

- `dist\\Clipboard 1.0.0\\Clipboard.exe`

分发时请发送整个 `dist\\Clipboard 1.0.0` 文件夹，而不是只发单个 `Clipboard.exe`。

## 数据目录

- 数据库：`%APPDATA%\\Clipboard\\history.db`
- 图片缓存：`%APPDATA%\\Clipboard\\images\\`
