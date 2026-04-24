# macOS adapter notes

The macOS adapter is implemented around PyObjC `NSPasteboard` plus a small
`rumps` menu-bar helper process. It is designed for macOS 12+ and still needs
final acceptance on a real Mac before publishing a release asset.

Current coverage:

- text, HTML, RTF, screenshot image, image-file, multi-image, and file-url reads
- text, HTML, RTF, single-image, multi-image, and mixed text/image writes
- same-folder portable `data/`
- menu-bar commands through a helper subprocess and Unix socket
- single-instance lock plus wake-up socket command
- LaunchAgent startup registration

Distribution notes:

- `pyinstaller clipboard_macos.spec` must be run on macOS.
- The generated `.app` is not signed or notarized.
- Release zips must not include `data/`, `history.db`, or image cache files.
