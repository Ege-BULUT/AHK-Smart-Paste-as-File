#Requires AutoHotkey v2.0
#SingleInstance Force

; =========================================================
; SMART CLIPBOARD FILE PASTER
;
; Hotkey:
;   ALT + V
;
; What it does:
;   Inspects clipboard content, creates a real file, then:
;
;   - If Windows Explorer is focused:
;       Saves the file directly into the active Explorer folder.
;
;   - If Windows Explorer is NOT focused:
;       Saves the file into Documents\Clipboard,
;       puts that file into clipboard as a real file using CF_HDROP,
;       then sends Ctrl+V to the focused app.
;
; Supported outputs:
;   - Image       -> .png
;   - Base64 blob -> detected file type: .pdf/.png/.jpg/.gif/.webp/.zip/.json/.txt/.bin
;   - Table       -> .xlsx, fallback .csv
;   - JSON        -> .json
;   - Python      -> .py
;   - AutoHotkey  -> .ahk
;   - Markdown    -> .md
;   - HTML        -> .html
;   - JavaScript  -> .js
;   - TypeScript  -> .ts
;   - Shell       -> .sh
;   - YAML        -> .yaml
;   - Plain text  -> .txt
;
; Notes:
;   - XLSX requires Microsoft Excel installed.
;   - If Excel fails, table content is saved as CSV.
;   - Image saving first tries the old stable GDI+ HBITMAP method,
;     then falls back to PowerShell clipboard image extraction.
;
; =========================================================

global DEFAULT_DIR := A_MyDocuments "\Clipboard"

if !DirExist(DEFAULT_DIR)
    DirCreate(DEFAULT_DIR)

; =========================================================
; HOTKEY
; =========================================================

!v::
{
    explorerFocused := IsExplorerFocused()
    targetDir := GetTargetDirectory()

    if !DirExist(targetDir)
        DirCreate(targetDir)

    timestamp := FormatTime(, "yyyy-MM-dd_HHmmss")

    ; -----------------------------------------------------
    ; IMAGE
    ; -----------------------------------------------------
    if ClipboardHasImage()
    {
        filePath := GetUniqueFileName(targetDir, timestamp, "png")

        if SaveClipboardImageAsPNG(filePath)
        {
            FinishFileAction(filePath, explorerFocused)
            return
        }

        MsgBox "Failed to save clipboard image."
        return
    }

    ; -----------------------------------------------------
    ; TEXT / BASE64 / CODE / TABLE
    ; -----------------------------------------------------
    clip := A_Clipboard

    if Trim(clip) = ""
    {
        MsgBox "Clipboard boş veya desteklenmeyen format."
        return
    }

    type := DetectTextType(clip)

    switch type
    {
        case "base64":
        {
            filePath := SaveBase64ToFile(clip, targetDir, timestamp)

            if filePath
            {
                FinishFileAction(filePath, explorerFocused)
                return
            }

            MsgBox "Failed to decode base64."
            return
        }

        case "table":
        {
            xlsxPath := GetUniqueFileName(targetDir, timestamp, "xlsx")

            if SaveTextAsExcel(clip, xlsxPath)
            {
                FinishFileAction(xlsxPath, explorerFocused)
                return
            }

            csvPath := GetUniqueFileName(targetDir, timestamp, "csv")

            if SaveTextAsCsv(clip, csvPath)
            {
                FinishFileAction(csvPath, explorerFocused)
                TrayTip "SmartPaste", "XLSX failed, saved CSV instead."
                return
            }

            MsgBox "Failed to generate table file."
            return
        }

        case "json":
            ext := "json"

        case "python":
            ext := "py"

        case "ahk":
            ext := "ahk"

        case "markdown":
            ext := "md"

        case "html":
            ext := "html"

        case "javascript":
            ext := "js"

        case "typescript":
            ext := "ts"

        case "shell":
            ext := "sh"

        case "yaml":
            ext := "yaml"

        default:
            ext := "txt"
    }

    filePath := GetUniqueFileName(targetDir, timestamp, ext)

    try
    {
        FileAppend clip, filePath, "UTF-8"
    }
    catch as e
    {
        MsgBox "Failed to write file:`n`n" e.Message
        return
    }

    FinishFileAction(filePath, explorerFocused)
}

; =========================================================
; FINAL ACTION
; =========================================================

FinishFileAction(filePath, explorerFocused)
{
    if explorerFocused
    {
        try Send "{F5}"
        TrayTip "SmartPaste", "Saved: " filePath
        return
    }

    if PasteFile(filePath)
        TrayTip "SmartPaste", "Pasted file: " filePath
    else
        MsgBox "File was created but could not be pasted:`n`n" filePath
}

; =========================================================
; TARGET DIRECTORY
; =========================================================

IsExplorerFocused()
{
    try
    {
        hwnd := WinActive("A")
        class := WinGetClass(hwnd)
        return class = "CabinetWClass" || class = "ExploreWClass"
    }
    catch
    {
        return false
    }
}

GetTargetDirectory()
{
    if IsExplorerFocused()
    {
        path := GetActiveExplorerPath()

        if path && DirExist(path)
            return path
    }

    if !DirExist(DEFAULT_DIR)
        DirCreate(DEFAULT_DIR)

    return DEFAULT_DIR
}

GetActiveExplorerPath()
{
    try
    {
        activeHwnd := WinActive("A")

        for window in ComObject("Shell.Application").Windows
        {
            try
            {
                if window.HWND = activeHwnd
                    return window.Document.Folder.Self.Path
            }
        }
    }

    return ""
}

GetUniqueFileName(folder, baseName, ext)
{
    filePath := folder "\" baseName "." ext
    counter := 1

    while FileExist(filePath)
    {
        filePath := folder "\" baseName " (" counter ")." ext
        counter++
    }

    return filePath
}

; =========================================================
; TYPE DETECTION
; =========================================================

DetectTextType(text)
{
    trimmed := Trim(text)

    if IsLikelyBase64(trimmed)
        return "base64"

    if IsLikelyJson(trimmed)
        return "json"

    if IsLikelyMarkdown(trimmed)
        return "markdown"

    if IsLikelyHTML(trimmed)
        return "html"

    if IsLikelyTypeScript(trimmed)
        return "typescript"

    if IsLikelyJavaScript(trimmed)
        return "javascript"

    if IsLikelyAHK(trimmed)
        return "ahk"

    if IsLikelyShell(trimmed)
        return "shell"

    if IsLikelyYAML(trimmed)
        return "yaml"

    if IsLikelyPython(trimmed)
        return "python"

    if IsLikelyTable(trimmed)
        return "table"

    return "text"
}

; =========================================================
; DETECTORS
; =========================================================

IsLikelyJson(text)
{
    text := Trim(text)

    first := SubStr(text, 1, 1)
    last := SubStr(text, -1)

    if !(
        (first = "{" && last = "}")
        ||
        (first = "[" && last = "]")
    )
        return false

    ; Object-like JSON
    if InStr(text, Chr(34) ":")
        return true

    ; Common JSON literals
    if InStr(text, "null") || InStr(text, "true") || InStr(text, "false")
        return true

    ; Array-like JSON without quote escaping problems
    if first = "["
        return true

    return false
}

IsLikelyMarkdown(text)
{
    if RegExMatch(text, "m)^# ")
        return true

    if RegExMatch(text, "m)^## ")
        return true

    if InStr(text, Chr(96) Chr(96) Chr(96))
        return true

    if RegExMatch(text, "\[.*\]\(.*\)")
        return true

    if RegExMatch(text, "m)^\- ")
        return true

    return false
}

IsLikelyHTML(text)
{
    if InStr(text, "<html")
        return true

    if InStr(text, "<body")
        return true

    if InStr(text, "<div")
        return true

    if InStr(text, "<span")
        return true

    if InStr(text, "<script")
        return true

    if RegExMatch(text, "i)<[a-z][a-z0-9\-]*(\s[^>]*)?>")
        return true

    return false
}

IsLikelyJavaScript(text)
{
    score := 0

    patterns := [
        "console.log("
        ,"function("
        ,"function "
        ,"=>"
        ,"const "
        ,"let "
        ,"var "
        ,"document."
        ,"window."
        ,"require("
        ,"module.exports"
        ,"export default"
        ,"import "
    ]

    for _, p in patterns
    {
        if InStr(text, p)
            score++
    }

    return score >= 2
}

IsLikelyTypeScript(text)
{
    score := 0

    patterns := [
        "interface "
        ,"type "
        ,": string"
        ,": number"
        ,": boolean"
        ,"implements "
        ,"enum "
        ,"<T>"
        ,"as const"
        ,"React.FC"
    ]

    for _, p in patterns
    {
        if InStr(text, p)
            score++
    }

    return score >= 2
}

IsLikelyAHK(text)
{
    score := 0

    patterns := [
        "#Requires"
        ,"#SingleInstance"
        ,"A_Clipboard"
        ,"Send("
        ,"Send "
        ,"MsgBox"
        ,"WinActivate"
        ,"WinExist"
        ,"ControlSend"
        ,"CoordMode"
        ,"SetTimer"
        ,"Hotkey"
        ,"Click"
        ,"MouseMove"
        ,"DllCall("
        ,"ComObject("
        ,"TrayTip"
    ]

    for _, p in patterns
    {
        if InStr(text, p)
            score++
    }

    if RegExMatch(text, "m)^\s*[#!^+]*[A-Za-z0-9]+::")
        score += 2

    return score >= 2
}

IsLikelyShell(text)
{
    if InStr(text, "#!/bin/bash")
        return true

    if InStr(text, "#!/bin/sh")
        return true

    if RegExMatch(text, "m)^sudo ")
        return true

    if RegExMatch(text, "m)^apt ")
        return true

    if RegExMatch(text, "m)^echo ")
        return true

    if RegExMatch(text, "m)^cd ")
        return true

    if RegExMatch(text, "m)^chmod ")
        return true

    return false
}

IsLikelyYAML(text)
{
    if InStr(text, "{") || InStr(text, "}")
        return false

    yamlLines := 0
    lines := StrSplit(text, "`n")

    for _, line in lines
    {
        line := Trim(line, "`r")

        if RegExMatch(line, "^\s*[A-Za-z0-9_\-]+\s*:\s*.+$")
            yamlLines++

        if RegExMatch(line, "^\s*-\s+")
            yamlLines++
    }

    return yamlLines >= 2
}

IsLikelyPython(text)
{
    score := 0

    patterns := [
        "def "
        ,"class "
        ,"import "
        ,"from "
        ,"if __name__"
        ,"print("
        ,"self."
        ,"try:"
        ,"except"
        ,"elif "
        ,"async def"
        ,"await "
        ,"lambda "
        ,"return "
    ]

    for _, p in patterns
    {
        if InStr(text, p)
            score++
    }

    if RegExMatch(text, "m)^(def|class|if|for|while|try|with).+:$")
        score++

    return score >= 2
}

IsLikelyTable(text)
{
    lines := StrSplit(Trim(text), "`n")

    if lines.Length < 2
        return false

    tabRows := 0
    commaRows := 0

    for _, line in lines
    {
        line := Trim(line, "`r")

        if line = ""
            continue

        if InStr(line, "`t")
            tabRows++

        if CountChar(line, ",") >= 1
            commaRows++
    }

    if tabRows >= 2
        return true

    if commaRows >= 2
        return true

    return false
}

CountChar(text, char)
{
    return StrLen(text) - StrLen(StrReplace(text, char))
}

IsLikelyBase64(text)
{
    text := Trim(text)

    ; Remove Data URI prefix if present.
    text := RegExReplace(text, "^data:.*?;base64,", "")
    cleaned := RegExReplace(text, "\s", "")

    if StrLen(cleaned) < 128
        return false

    if Mod(StrLen(cleaned), 4) != 0
        return false

    if !RegExMatch(cleaned, "^[A-Za-z0-9+/=]+$")
        return false

    return true
}

; =========================================================
; BASE64
; =========================================================

SaveBase64ToFile(base64Text, targetDir, timestamp)
{
    try
    {
        ext := "bin"

        if RegExMatch(base64Text, "^data:(.*?);base64,", &m)
        {
            ext := MimeToExtension(m[1])
            base64Text := RegExReplace(base64Text, "^data:.*?;base64,", "")
        }

        bytes := Base64Decode(base64Text)

        if !bytes || bytes.Size = 0
            return false

        if ext = "bin"
            ext := DetectBinaryExtension(bytes)

        filePath := GetUniqueFileName(targetDir, timestamp, ext)

        f := FileOpen(filePath, "w")
        f.RawWrite(bytes, bytes.Size)
        f.Close()

        return filePath
    }
    catch
    {
        return false
    }
}

MimeToExtension(mime)
{
    switch mime
    {
        case "application/pdf":
            return "pdf"

        case "image/png":
            return "png"

        case "image/jpeg":
            return "jpg"

        case "image/webp":
            return "webp"

        case "image/gif":
            return "gif"

        case "text/plain":
            return "txt"

        case "application/json":
            return "json"

        case "application/zip":
            return "zip"

        case "text/html":
            return "html"

        case "text/markdown":
            return "md"

        default:
            return "bin"
    }
}

DetectBinaryExtension(bytes)
{
    hex := ""
    max := Min(16, bytes.Size)

    loop max
        hex .= Format("{:02X}", NumGet(bytes, A_Index - 1, "UChar"))

    if InStr(hex, "25504446")
        return "pdf"

    if InStr(hex, "89504E47")
        return "png"

    if InStr(hex, "FFD8FF")
        return "jpg"

    if InStr(hex, "504B0304")
        return "zip"

    if InStr(hex, "47494638")
        return "gif"

    if InStr(hex, "52494646")
        return "webp"

    return "bin"
}

Base64Decode(b64)
{
    b64 := RegExReplace(b64, "\s", "")

    size := 0

    ok := DllCall(
        "Crypt32.dll\CryptStringToBinaryW"
        ,"Str", b64
        ,"UInt", 0
        ,"UInt", 1
        ,"Ptr", 0
        ,"UInt*", &size
        ,"Ptr", 0
        ,"Ptr", 0
    )

    if !ok || size = 0
        return false

    buf := Buffer(size, 0)

    ok := DllCall(
        "Crypt32.dll\CryptStringToBinaryW"
        ,"Str", b64
        ,"UInt", 0
        ,"UInt", 1
        ,"Ptr", buf
        ,"UInt*", &size
        ,"Ptr", 0
        ,"Ptr", 0
    )

    if !ok
        return false

    return buf
}

; =========================================================
; XLSX / CSV
; =========================================================

SaveTextAsExcel(text, filePath)
{
    excel := ""
    wb := ""

    try
    {
        excel := ComObject("Excel.Application")
        excel.Visible := false
        excel.DisplayAlerts := false

        wb := excel.Workbooks.Add()
        ws := wb.Worksheets(1)

        rows := StrSplit(text, "`n")

        r := 1

        for _, rowText in rows
        {
            rowText := Trim(rowText, "`r")

            if InStr(rowText, "`t")
                cols := StrSplit(rowText, "`t")
            else
                cols := ParseCsvLine(rowText)

            c := 1

            for _, value in cols
            {
                ws.Cells(r, c).Value := value
                c++
            }

            r++
        }

        ws.Columns.AutoFit()

        ; 51 = xlOpenXMLWorkbook = .xlsx
        wb.SaveAs(filePath, 51)

        wb.Close(false)
        excel.Quit()

        return true
    }
    catch
    {
        try
        {
            if wb
                wb.Close(false)
        }

        try
        {
            if excel
                excel.Quit()
        }

        return false
    }
}

SaveTextAsCsv(text, filePath)
{
    try
    {
        rows := StrSplit(text, "`n")
        out := ""

        for _, rowText in rows
        {
            rowText := Trim(rowText, "`r")

            if InStr(rowText, "`t")
            {
                cols := StrSplit(rowText, "`t")
                out .= CsvJoin(cols) "`r`n"
            }
            else
            {
                out .= rowText "`r`n"
            }
        }

        FileAppend out, filePath, "UTF-8"
        return true
    }
    catch
    {
        return false
    }
}

CsvJoin(cols)
{
    out := ""

    for i, value in cols
    {
        escaped := StrReplace(value, '"', '""')
        out .= '"' escaped '"'

        if i < cols.Length
            out .= ","
    }

    return out
}

ParseCsvLine(line)
{
    cols := []
    current := ""
    inQuotes := false
    i := 1

    while i <= StrLen(line)
    {
        ch := SubStr(line, i, 1)

        if ch = '"'
        {
            nextCh := SubStr(line, i + 1, 1)

            if inQuotes && nextCh = '"'
            {
                current .= '"'
                i += 2
                continue
            }

            inQuotes := !inQuotes
            i++
            continue
        }

        if ch = "," && !inQuotes
        {
            cols.Push(current)
            current := ""
            i++
            continue
        }

        current .= ch
        i++
    }

    cols.Push(current)
    return cols
}

; =========================================================
; IMAGE SUPPORT
; =========================================================

ClipboardHasImage()
{
    ; CF_BITMAP = 2, CF_DIB = 8, CF_DIBV5 = 17.
    return DllCall("IsClipboardFormatAvailable", "UInt", 2)
        || DllCall("IsClipboardFormatAvailable", "UInt", 8)
        || DllCall("IsClipboardFormatAvailable", "UInt", 17)
}

SaveClipboardImageAsPNG(filePath)
{
    ; Primary: old stable CF_BITMAP + GDI+ approach.
    if DllCall("IsClipboardFormatAvailable", "UInt", 2)
    {
        if DllCall("OpenClipboard", "Ptr", 0)
        {
            hBitmap := DllCall("GetClipboardData", "UInt", 2, "Ptr")
            DllCall("CloseClipboard")

            if hBitmap
            {
                try
                {
                    if HBitmapToPngFile(hBitmap, filePath) && FileExist(filePath)
                        return true
                }
                catch
                {
                }
            }
        }
    }

    ; Fallback: PowerShell clipboard image extraction.
    return SaveClipboardImageAsPNG_PowerShell(filePath)
}

HBitmapToPngFile(hBitmap, dest)
{
    hMod := 0

    if !DllCall("GetModuleHandle", "Str", "gdiplus.dll", "Ptr")
        hMod := DllCall("LoadLibrary", "Str", "gdiplus.dll", "Ptr")

    si := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
    NumPut("UInt", 1, si, 0)

    pToken := 0

    status := DllCall(
        "gdiplus\GdiplusStartup",
        "UPtr*", &pToken,
        "Ptr", si,
        "Ptr", 0
    )

    if status != 0
        return false

    pBitmap := 0

    status := DllCall(
        "gdiplus\GdipCreateBitmapFromHBITMAP",
        "Ptr", hBitmap,
        "Ptr", 0,
        "UPtr*", &pBitmap
    )

    if status != 0 || !pBitmap
    {
        DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)

        if hMod
            DllCall("FreeLibrary", "Ptr", hMod)

        return false
    }

    CLSID := Buffer(16, 0)

    DllCall(
        "ole32\CLSIDFromString",
        "WStr", "{557CF406-1A04-11D3-9A73-0000F81EF32E}",
        "Ptr", CLSID
    )

    status := DllCall(
        "gdiplus\GdipSaveImageToFile",
        "Ptr", pBitmap,
        "WStr", dest,
        "Ptr", CLSID,
        "Ptr", 0
    )

    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)

    if hMod
        DllCall("FreeLibrary", "Ptr", hMod)

    return status = 0
}

SaveClipboardImageAsPNG_PowerShell(filePath)
{
    psFile := A_Temp "\smartpaste_save_clipboard_image.ps1"
    escapedPath := StrReplace(filePath, "'", "''")

    psScript := "$ErrorActionPreference = 'Stop'`r`n"
    psScript .= "Add-Type -AssemblyName System.Windows.Forms`r`n"
    psScript .= "Add-Type -AssemblyName System.Drawing`r`n"
    psScript .= "$img = [System.Windows.Forms.Clipboard]::GetImage()`r`n"
    psScript .= "if ($null -eq $img) { exit 2 }`r`n"
    psScript .= "$img.Save('" escapedPath "', [System.Drawing.Imaging.ImageFormat]::Png)`r`n"

    try FileDelete(psFile)
    FileAppend psScript, psFile, "UTF-8"

    try
    {
        exitCode := RunWait(
            'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' psFile '"',
            ,
            "Hide"
        )

        return exitCode = 0 && FileExist(filePath)
    }
    catch
    {
        return false
    }
}

; =========================================================
; REAL FILE CLIPBOARD / PASTE
; =========================================================

PasteFile(filePath)
{
    if !SetFilesToClipboard([filePath])
        return false

    Sleep 120
    Send "^v"
    return true
}

SetFilesToClipboard(paths)
{
    if paths.Length = 0
        return false

    ; DROPFILES structure:
    ; 20 bytes header + UTF-16 file list separated by NUL + double-NUL terminator.
    totalChars := 1

    for _, path in paths
        totalChars += StrLen(path) + 1

    size := 20 + totalChars * 2

    hDrop := DllCall("GlobalAlloc", "UInt", 0x42, "Ptr", size, "Ptr") ; GMEM_MOVEABLE | GMEM_ZEROINIT

    if !hDrop
        return false

    pDrop := DllCall("GlobalLock", "Ptr", hDrop, "Ptr")

    if !pDrop
    {
        DllCall("GlobalFree", "Ptr", hDrop)
        return false
    }

    NumPut("UInt", 20, pDrop, 0) ; pFiles offset
    NumPut("Int", 0, pDrop, 4)
    NumPut("Int", 0, pDrop, 8)
    NumPut("Int", 0, pDrop, 12)
    NumPut("Int", 1, pDrop, 16) ; fWide = TRUE

    offset := 20

    for _, path in paths
    {
        StrPut(path, pDrop + offset, StrLen(path) + 1, "UTF-16")
        offset += (StrLen(path) + 1) * 2
    }

    NumPut("UShort", 0, pDrop, offset)

    DllCall("GlobalUnlock", "Ptr", hDrop)

    if !OpenClipboardSafe()
    {
        DllCall("GlobalFree", "Ptr", hDrop)
        return false
    }

    DllCall("EmptyClipboard")

    ok := DllCall("SetClipboardData", "UInt", 15, "Ptr", hDrop, "Ptr") ; CF_HDROP = 15

    DllCall("CloseClipboard")

    if !ok
    {
        DllCall("GlobalFree", "Ptr", hDrop)
        return false
    }

    ; On success, Windows owns hDrop. Do not free it.
    return true
}

OpenClipboardSafe()
{
    loop 10
    {
        if DllCall("OpenClipboard", "Ptr", 0)
            return true

        Sleep 50
    }

    return false
}
