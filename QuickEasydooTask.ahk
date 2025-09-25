#Requires AutoHotkey v2.0
; Make rendering crisp on Win10/11 multi-DPI setups
try DllCall("user32\SetProcessDpiAwarenessContext", "ptr", -4) ; PER_MONITOR_AWARE_V2
global gBar := 0, gEdit := 0, EditBrush := 0

; ====== CONFIG ======
API_URL := "https://api.easydoo.com/api/integration/CreateOrUpdateEntity"

API_KEY := EnvGet("EASYDOO_API_KEY")
if !API_KEY {
    FlashTip("Missing EASYDOO_API_KEY env var.", 4000)
    return
}
DEBUG := 0   ; 1 = show payload/response MsgBoxes, 0 = silent

; ====================

; Hotkey: Ctrl+Alt+T -> show input bar
^!k::ShowInputBar()
#e::Run('"' EnvGet("LocalAppData") '\VoidStar\FilePilot\FPilot.exe" "C:"')

; Help hotkey: Ctrl+Win+Alt+H -> show all shortcuts
^#!h::ShowHelpDialog()



ShowInputBar() {
    global gBar, gEdit, EditBrush

    if !gBar {
        gBar := Gui("+AlwaysOnTop -Caption +ToolWindow", "Quick POST")
        gBar.SetFont("s20", "FiraCode Nerd Font")

        ApplyWin11Effects(gBar.Hwnd, Map(
            "Dark", true,
            "Corners", 2,
            "Backdrop", 2,
            "Shadow", true
        ))

        gBar.BackColor := "1E1E1E"
        gBar.Add("Text", "x0 y0 w1080 h120 Background2D2D30 -TabStop")
        gBar.Add("Text", "x24 y26 w1032 h54 Background404040 -TabStop")

        gEdit := gBar.Add("Edit", "x28 y30 w1024 h46 -Border -E0x200")
        gEdit.SetFont("s20", "FiraCode Nerd Font")

        ; Always-dark edit box
        OnMessage(0x0133, EditColorHandler)
        OnMessage(0x0138, EditColorHandler)
        if !EditBrush
            EditBrush := DllCall("gdi32\CreateSolidBrush", "uint", 0x2D2D30, "ptr")

        try DllCall("uxtheme\SetWindowTheme", "ptr", gEdit.Hwnd, "wstr", "DarkMode_Explorer", "ptr", 0)

        gEdit.OnEvent("Focus", (*) => Send("^a"))
        SetCueBanner(gEdit.Hwnd, "Enter the name of the easydoo-workitem…")
        SetEditMargins(gEdit.Hwnd, 14, 14)

        gBar.Add("Text", "x24 y84 w1032 h1 Background404040 -TabStop")
        btn := gBar.Add("Button", "Default w0 h0"), btn.OnEvent("Click", (*) => Submit())
        gBar.OnEvent("Escape", (*) => (gBar.Hide(), ToolTip()))
    }

    gEdit.Value := ""

    ; --- Precise centering with DPI-correct size ---
    gBar.Show("Hide AutoSize")
    gBar.GetPos(, , &w, &h)

    MouseGetPos &mx, &my
    idx := MonitorGetPrimary()
    Loop MonitorGetCount() {
        MonitorGet(A_Index, &L, &T, &R, &B)
        if (mx >= L && mx < R && my >= T && my < B) {
            idx := A_Index
            break
        }
    }
    MonitorGetWorkArea(idx, &L, &T, &R, &B)

    dpi := DllCall("user32\GetDpiForWindow", "ptr", gBar.Hwnd, "uint")
    scale := dpi / 96.0
    pw := Round(w * scale), ph := Round(h * scale)

    x := L + ((R - L) - pw) // 2
    y := T + ((B - T) - ph) // 2
    gBar.Show(Format("x{} y{}", x, y))
    gEdit.Focus()
}

; Keeps edit control dark with white text
EditColorHandler(wParam, lParam, msg, hwnd) {
    global gEdit, EditBrush
    if (lParam = gEdit.Hwnd) {
        DllCall("gdi32\SetTextColor", "ptr", wParam, "uint", 0xFFFFFF)
        DllCall("gdi32\SetBkColor", "ptr", wParam, "uint", 0x2D2D30)
        return EditBrush
    }
}

#HotIf gBar && WinActive("ahk_id " gBar.Hwnd)
Enter::Submit()
Esc::(gBar.Hide(), ToolTip())
#HotIf

Submit() {
    global gBar, gEdit, API_URL, API_KEY
    title := Trim(gEdit.Value)
    if (title = "") {
        FlashTip("Please type a title first.")
        return
    }
    gBar.Hide()
    payload := BuildJsonWithTitle(title)

    if (DEBUG)
        MsgBox("Payload to send:`n`n" payload)

    try {
        WHR := ComObject("WinHttp.WinHttpRequest.5.1")
        WHR.Open("POST", API_URL, true)
        WHR.SetRequestHeader("Content-Type", "application/json")
        WHR.SetRequestHeader("x-api-key", API_KEY)
        WHR.SetTimeouts(30000, 30000, 30000, 30000)

        WHR.Send(payload)
        WHR.WaitForResponse()
        status := WHR.Status
        resp := WHR.ResponseText

        if (DEBUG)
            MsgBox("HTTP Status: " status "`n`nResponse:`n`n" resp)

        if (status >= 200 && status < 300)
            FlashTip("Sent ✓ (" status ")", 1800)
        else
            FlashTip("Error " status ": " Shorten(resp), 4000)
    } catch as e {
        if (DEBUG)
            MsgBox("Request failed:`n" e.Message)
        FlashTip("Request failed: " e.Message, 4000)
    }
}

; --- Helpers ---
BuildJsonWithTitle(title) {
    q := Chr(34)
    return "{"
        . q "entity" q ": {"
            . q "$type" q ":" q "WorkItem" q ","
            . q "assignedTo" q ": {"
                . q "id" q ":" q "0cbfe7f3-a091-4ca2-8dda-85bb43f73c31" q
            . "},"
            . q "workItemType" q ": {"
                . q "id" q ":" q "24e77918-184b-4bfb-9ba6-49d9dde68748" q
            . "},"
            . q "state" q ": {"
                . q "id" q ":" q "2e50bade-1641-4af2-b978-ea1e0f2e84b4" q
            . "},"
            . q "sharedWorkFolders" q ": {"
                . q "addOrUpdateItems" q ": ["
                    . "{"
                        . q "id" q ":" q "2467b91b-4523-457e-a52b-65b0a9e8bb2b" q
                    . "}"
                . "]"
            . "},"
            . q "customFields" q ": ["
                . "{"
                    . q "value" q ":" q "" q ","
                    . q "fieldDefinition" q ": {"
                        . q "id" q ":" q "42f82244-0ee7-89c8-8cd4-20f4fb80b71e" q
                    . "}"
                . "}"
            . "],"
            . q "name" q ":" q JsonEscape(title) q ","
            . q "IsVisible" q ": true"
        . "},"
        . q "workspaceId" q ":" q "3ef492ff-3165-4cbf-a7cf-2003234703ca" q
    . "}"
}

JsonEscape(str) {
    q := Chr(34), b := Chr(92), out := ""
    for c in StrSplit(str, "") {
        code := Ord(c)
        if (c = q)
            out .= b . q
        else if (c = b)
            out .= b . b
        else if (code = 8)
            out .= b "b"
        else if (code = 9)
            out .= b "t"
        else if (code = 10)
            out .= b "n"
        else if (code = 12)
            out .= b "f"
        else if (code = 13)
            out .= b "r"
        else if (code < 32)
            out .= Format("\u{:04X}", code)
        else
            out .= c
    }
    return out
}

Shorten(s, max := 160) {
    s := Trim(s)
    return (StrLen(s) > max) ? SubStr(s, 1, max-1) "…" : s
}

FlashTip(text, ms := 2000) {
    ToolTip(text)
    SetTimer(() => ToolTip(), -ms)
}

SetCueBanner(hEdit, text) {
    DllCall("user32\SendMessageW", "ptr", hEdit, "uint", 0x1501, "ptr", 1, "wstr", text)
}

SetEditMargins(hEdit, leftPx := 8, rightPx := 8) {
    lParam := leftPx | (rightPx << 16)
    DllCall("user32\SendMessageW", "ptr", hEdit, "uint", 0x00D3, "ptr", 0x3, "ptr", lParam)
}

SetEditColors(hEdit) {
    DllCall("user32\SendMessageW", "ptr", hEdit, "uint", 0x00C4, "ptr", 0, "ptr", 0x2D2D30)
    try DllCall("uxtheme\SetWindowTheme", "ptr", hEdit, "wstr", "DarkMode_Explorer", "ptr", 0)
    try {
        val := 1
        DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hEdit, "uint", 20, "int*", val, "uint", 4)
    }
}

ApplyWin11Effects(hwnd, opts := Map(
    "Dark", true,
    "Corners", 2,
    "Backdrop", 2,
    "Shadow", true
)) {
    if opts["Dark"] {
        val := 1
        DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "uint", 20, "int*", val, "uint", 4)
    }
    if opts.Has("Corners") {
        corner := Integer(opts["Corners"])
        DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "uint", 33, "int*", corner, "uint", 4)
    }
    if opts.Has("Backdrop") {
        bt := Integer(opts["Backdrop"])
        DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "uint", 38, "int*", bt, "uint", 4)
    }
    if opts["Shadow"] {
        margins := Buffer(16, 0)
        NumPut("int", 1, margins, 0), NumPut("int", 1, margins, 4)
        NumPut("int", 1, margins, 8), NumPut("int", 1, margins, 12)
        DllCall("dwmapi\DwmExtendFrameIntoClientArea", "ptr", hwnd, "ptr", margins)
    }
}

; Hotkey: Ctrl+Alt+Win+O
^!#o:: {
    created := 0
    skipped := 0
    failed  := 0

    sel := GetSelectedPaths()
    if sel.Length = 0 {
        MsgBox "Kein Element in Datei-Explorer ausgewählt.", "Hinweis", "Iconi"
        return
    }

    ; Find 7z.exe (try common locations, then PATH)
    candidates := [
        A_ProgramFiles "\7-Zip\7z.exe",
        A_ProgramFiles "\7-Zip-Zstandard\7z.exe",
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe",
        "7z.exe"
    ]
    sevenZip := ""
    for p in candidates {
        if FileExist(p) {
            sevenZip := p
            break
        }
    }
    if sevenZip = "" {
        MsgBox "7-Zip (7z.exe) wurde nicht gefunden. Bitte 7-Zip installieren " 
             . "oder den Pfad in der Liste 'candidates' anpassen.", "Fehler", "Iconx"
        return
    }

    fmt := EnvGet("ZIP_PW_FORMAT")
    if (fmt = "") {
        MsgBox "Umgebungsvariable 'ZIP_PW_FORMAT' ist nicht gesetzt!", "Fehler", "Iconx"
        return
    }

    ; Datumstokens einsetzen: {YYYY}, {YY}, {MM}, {M}
    yyyy := FormatTime(, "yyyy")
    mm   := FormatTime(, "MM")
    password := fmt
    password := StrReplace(password, "{YYYY}", yyyy)
    password := StrReplace(password, "{YY}",   FormatTime(, "yy"))
    password := StrReplace(password, "{MM}",   mm)
    password := StrReplace(password, "{M}",    FormatTime(, "M"))

    for filePath in sel {
        ; Skip .zip sources themselves
        if StrLower(RegExReplace(filePath, ".*\.")) = "zip" {
            skipped++
            continue
        }

        dir := SplitPath(filePath, &name, &dirPath, &ext, &nameNoExt)
        zipPath := dirPath "\" nameNoExt ".zip"

        if FileExist(zipPath) {
            skipped++
            continue
        }

        ; Build 7z command (quote password to allow spaces/specials)
        ; 7z a -tzip -p"PASSWORD" -mem=AES256 -y -- "zipPath" "filePath"
        cmd := Format('"{1}" a -tzip -p"{2}" -mem=AES256 -y -- "{3}" "{4}"'
                    , sevenZip, password, zipPath, filePath)

        try {
            ret := RunWait(cmd, , "Hide")
            if FileExist(zipPath) {
                created++
                ; Delete original only if it is a file, not a folder
                if InStr(FileGetAttrib(filePath), "D") = 0 {
                    FileDelete filePath
                }
            } else {
                failed++
            }
        } catch {
            failed++
        }
    }

    MsgBox Format("Fertig.`nErstellt: {1}`nUebersprungen: {2}`nFehlgeschlagen: {3}"
                , created, skipped, failed), "Ergebnis", "Iconi"
}

; -------- Helpers --------

GetSelectedPaths() {
    paths := []
    hwnd := WinActive("ahk_class CabinetWClass")      ; File Explorer
    if !hwnd
        hwnd := WinActive("ahk_class ExploreWClass")  ; (rare older class)
    if !hwnd
        return paths

    shell := ComObject("Shell.Application")
    for win in shell.Windows {
        try {
            if win.hwnd = hwnd {
                for item in win.Document.SelectedItems
                    paths.Push(item.Path)
                break
            }
        }
    }
    return paths
}

SplitPath(full, &name?, &dir?, &ext?, &nameNoExt?) {
    name := RegExReplace(full, ".+\\")
    dir  := SubStr(full, 1, StrLen(full) - StrLen(name) - (InStr(full, "\") ? 0 : 0))
    ext  := RegExReplace(name, ".*\.")
    nameNoExt := RegExReplace(name, "\.[^.]+$")
    return dir
}

; Show help dialog with all available hotkeys
ShowHelpDialog() {
    helpText := "Available Hotkeys:`n`n"
              . "Ctrl+Alt+K - Show Quick Task Input`n"
              . "   Creates a new EasyDoo work item with the entered title`n`n"
              . "Win+E - Open File Explorer`n"
              . "   Opens FilePilot file manager at C: drive`n`n"
              . "Ctrl+Win+Alt+O - ZIP Selected Files`n"
              . "   Creates password-protected ZIP files from selected items`n"
              . "   (Uses ZIP_PW_FORMAT environment variable for password)`n`n"
              . "Ctrl+Win+Alt+H - Show This Help`n"
              . "   Displays all available keyboard shortcuts`n`n"
              . "Press OK to close this help."
    
    MsgBox(helpText, "AutoHotkey Script - Available Shortcuts", 0)
}

