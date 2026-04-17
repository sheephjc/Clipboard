# macOS adapter notes

This adapter is source-prepared from Windows and still needs validation on a real Mac.

Expected v1 coverage:

- text, HTML, RTF, image, and file-url clipboard reads
- text, HTML, RTF, and PNG clipboard writes
- same-folder portable `data/`
- best-effort menu bar entry through `rumps`

Known follow-up work after first Mac run:

- verify `NSPasteboard` type constants across target macOS versions
- decide whether startup registration should use LaunchAgent or a user-facing setting only
- generate a proper `.icns` icon for `Clipboard.app`
- sign/notarize the app if distributing beyond local use

