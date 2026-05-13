## Author
Ege BULUT

---

## License

MIT License.

You are free to use, modify, and distribute this script.

---

## Installation

### 1. Install AutoHotkey v2

Download and install AutoHotkey v2 from:

```txt
https://www.autohotkey.com/
```

Make sure you install **AutoHotkey v2**, not v1.

### 2. Download the script

Save the script as:

```txt
SmartPaste.ahk
```

### 3. Run the script

Double-click:

```txt
SmartPaste.ahk
```

The script will run in the background.

### 4. Use the shortcut

Copy something, then press:

```txt
Alt + V
```

---

## Requirements

### Required

* Windows
* AutoHotkey v2

### Optional

* Microsoft Excel

Excel is only required for `.xlsx` table generation.

If Excel is not available or Excel COM automation fails, SmartPaste falls back to saving table data as `.csv`.

---

<img width="774" height="485" alt="image" src="https://github.com/user-attachments/assets/47c5af8e-51c5-4020-9f19-3cf6d52436a1" />


# SmartPaste

SmartPaste is an AutoHotkey v2 script that turns your clipboard content into a real file with a single shortcut.

Press `Alt + V`, and SmartPaste inspects the current clipboard content, detects its type, creates the correct file, and saves or pastes it automatically.

It is designed for developers, AI workflows, document processing, note-taking, and anyone who frequently copies code, JSON, images, tables, markdown, or base64 data and wants to turn them into actual files quickly.

---

## Features

- One shortcut: `Alt + V`
- Automatically detects clipboard content type
- Creates real files from clipboard content
- Saves directly into the active Windows Explorer folder
- Falls back to `Documents\Clipboard` when Explorer is not focused
- Supports real file pasting into apps using Windows `CF_HDROP`
- Preserves images as `.png`
- Supports many developer-friendly formats
- Supports table-to-Excel export
- Supports base64 decoding into real files

---

## Supported Clipboard Types

| Clipboard Content | Output File |
|---|---|
| Plain text | `.txt` |
| JSON | `.json` |
| Python code | `.py` |
| AutoHotkey code | `.ahk` |
| Markdown | `.md` |
| HTML | `.html` |
| JavaScript | `.js` |
| TypeScript | `.ts` |
| Shell script | `.sh` |
| YAML | `.yaml` |
| Table / TSV / CSV-like text | `.xlsx` |
| Image | `.png` |
| Base64 PDF | `.pdf` |
| Base64 PNG | `.png` |
| Base64 JPEG | `.jpg` |
| Base64 GIF | `.gif` |
| Base64 WebP | `.webp` |
| Base64 ZIP | `.zip` |
| Unknown base64 binary | `.bin` |

---

## How It Works

SmartPaste behaves differently depending on the focused window.

### If Windows Explorer is focused

The file is created directly inside the currently open Explorer folder.

Example:

1. Copy an image
2. Open a folder in Windows Explorer
3. Press `Alt + V`
4. SmartPaste creates a `.png` file inside that folder

No extra paste operation is needed because the file is already created in the active folder.

### If another app is focused

The file is created inside:

```txt
Documents\Clipboard
````

Then SmartPaste puts that file into the clipboard as a real Windows file and sends `Ctrl + V` to the focused app.

This allows pasting generated files into apps such as:

* Discord
* Slack
* Teams
* Browsers
* ChatGPT
* Email clients
* Upload fields
* File managers
* Developer tools

---

## Examples

### Copy JSON

Clipboard:

```json
{
  "name": "SmartPaste",
  "type": "clipboard-tool",
  "enabled": true
}
```

Press `Alt + V`.

Output:

```txt
clipboard_2026-05-13_184233.json
```

---

### Copy Python Code

Clipboard:

```python
def hello():
    print("Hello from SmartPaste")
```

Press `Alt + V`.

Output:

```txt
clipboard_2026-05-13_184233.py
```

---

### Copy Markdown

Clipboard:

```md
# Project Notes

- Clipboard detection
- File generation
- AutoHotkey automation
```

Press `Alt + V`.

Output:

```txt
clipboard_2026-05-13_184233.md
```

---

### Copy an Image

Copy an image from a browser, screenshot tool, chat app, or image editor.

Press `Alt + V`.

Output:

```txt
clipboard_2026-05-13_184233.png
```

---

### Copy Table Data

Clipboard:

```txt
Name	Role	Language
Ege	Developer	Python
Alex	Designer	Figma
```

Press `Alt + V`.

Output:

```txt
clipboard_2026-05-13_184233.xlsx
```

If Excel is unavailable:

```txt
clipboard_2026-05-13_184233.csv
```

---

### Copy Base64

Clipboard:

```txt
JVBERi0xLjQKJc...
```

Press `Alt + V`.

SmartPaste decodes the content and detects the file type.

Output example:

```txt
clipboard_2026-05-13_184233.pdf
```

---

## File Naming

Generated files use timestamp-based names:

```txt
clipboard_YYYY-MM-DD_HHMMSS.ext
```

If a file with the same name already exists, SmartPaste creates a unique name:

```txt
clipboard_2026-05-13_184233 (1).txt
clipboard_2026-05-13_184233 (2).txt
```

---

## Default Output Folder

When Windows Explorer is not focused, SmartPaste saves files to:

```txt
Documents\Clipboard
```

The folder is created automatically if it does not already exist.

---

## Shortcut

Default shortcut:

```txt
Alt + V
```

In AutoHotkey syntax:

```ahk
!v::
```

You can change this by editing the hotkey line in the script.

Examples:

```ahk
^!v::   ; Ctrl + Alt + V
#+v::   ; Win + Shift + V
```

---

## Why Not Just Use Ctrl + V?

Normal `Ctrl + V` pastes clipboard content as-is.

SmartPaste does something different:

* Text becomes a `.txt` file
* JSON becomes a `.json` file
* Code becomes a source file
* Images become `.png`
* Tables become `.xlsx`
* Base64 becomes the original binary file

It turns clipboard content into a real file before pasting or saving it.

---

## Technical Details

SmartPaste uses several Windows-native techniques:

* AutoHotkey v2 for automation
* Windows clipboard format detection
* `CF_BITMAP`, `CF_DIB`, and `CF_DIBV5` checks for images
* GDI+ for image-to-PNG conversion
* PowerShell clipboard image extraction as fallback
* Excel COM automation for `.xlsx` generation
* `CF_HDROP` for real file clipboard paste
* Windows Shell COM for active Explorer folder detection
* `CryptStringToBinaryW` for base64 decoding

---

## Known Limitations

### XLSX support requires Excel

SmartPaste uses Excel COM automation to generate `.xlsx` files.

If Excel is not installed, table content is saved as `.csv`.

### Format detection is heuristic

The script uses practical detection rules rather than full parsers for every language.

For example, some text may be detected as Markdown, YAML, JavaScript, or plain text depending on its structure.

### Base64 detection requires reasonably large input

Very short base64 strings are ignored to avoid false positives.

### Some applications handle file paste differently

Most modern apps support file paste, but some apps may ignore pasted files depending on their input field or security restrictions.

---

## Recommended Use Cases

SmartPaste is especially useful for:

* Saving AI-generated code snippets
* Turning copied JSON responses into files
* Saving screenshots quickly
* Creating markdown notes from copied text
* Exporting copied tables to Excel
* Decoding base64 files
* Moving clipboard content into upload fields
* Building a cleaner developer workflow
* Quickly creating files inside the current Explorer folder

---

## Startup on Windows

To run SmartPaste automatically when Windows starts:

1. Press `Win + R`
2. Type:

```txt
shell:startup
```

3. Press Enter
4. Place a shortcut to `SmartPaste.ahk` in that folder

SmartPaste will now start automatically after login.

---

## Troubleshooting

### The script does not run

Make sure AutoHotkey v2 is installed.

If AutoHotkey v1 is installed instead, the script will not work correctly.

---

### Alt + V does nothing

Check that:

* The script is running
* The green AutoHotkey icon is visible in the system tray
* Another app is not capturing `Alt + V`
* Your clipboard is not empty

---

### Image saving fails

Try copying the image again.

Some apps expose images in unusual clipboard formats. SmartPaste first tries the native bitmap method and then falls back to PowerShell-based image extraction.

---

### XLSX generation fails

Install Microsoft Excel or use the CSV fallback.

If Excel is installed but XLSX still fails, Excel COM registration may be broken.

---

### File is created but not pasted into the app

Some applications do not accept file paste in every field.

Try focusing a file upload area, chat input, or message box that supports pasted files.

---

## Customization

You can customize:

* Hotkey
* Default output folder
* File naming format
* Detection rules
* Supported file extensions
* Table handling behavior
* Fallback behavior

The default output folder is defined here:

```ahk
global DEFAULT_DIR := A_MyDocuments "\Clipboard"
```

The hotkey is defined here:

```ahk
!v::
```
---

