#Requires AutoHotkey v2.0
; Make rendering crisp on Win10/11 multi-DPI setups
try DllCall("user32\SetProcessDpiAwarenessContext", "ptr", -4) ; PER_MONITOR_AWARE_V2


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
^!t::
{
    ShowInputBar()
}

; --- UI: minimal bar (borderless, always-on-top) ---
global gBar := 0, gEdit := 0

ShowInputBar() {
    global gBar, gEdit

    if !gBar {
        gBar := Gui("+AlwaysOnTop -Caption +ToolWindow", "Quick POST")

        ; --- Win11 look ---
        ApplyWin11Effects(gBar.Hwnd, Map(
            "Dark", true,        ; dark title bar
            "Corners", 2,        ; round
            "Backdrop", 2,       ; 2=Mica | 3=Transient(Acrylic-like) | 4=Tabbed(Mica Alt)
            "Shadow", true
        ))

        ; Use a near-transparent or themed background so the backdrop can “read” through
        ; Neutral surface (keeps text legible atop Mica/Acrylic)
        gBar.BackColor := "F3F4F6"

        ; Card
; Card
card := gBar.Add("Text", "x0 y0 w1080 h120 BackgroundTrans -TabStop")

        ; Edit
        gBar.SetFont("s20", "Segoe UI")
        gEdit := gBar.Add("Edit", "x24 y26 w1032 h54 -Wrap -VScroll -HScroll -Border")
       gEdit.BackColor := "F3F4F6"  ; or a dark tone if using dark mode

        gEdit.SetFont("s20 c202020")
        gEdit.OnEvent("Focus", (*) => Send("^a"))
        SetCueBanner(gEdit.Hwnd, "Enter the name of the easydoo-workitem…")
        SetEditMargins(gEdit.Hwnd, 14, 14)

        ; Divider (very soft)
        gBar.Add("Text", "x24 y84 w1032 h1 BackgroundCFCFCF -TabStop")

        ; Default hidden button & escape event
        btn := gBar.Add("Button", "Default w0 h0"), btn.OnEvent("Click", (*) => Submit())
        gBar.OnEvent("Escape", (*) => (gBar.Hide(), ToolTip()))
    }

    gEdit.Value := ""
    sw := A_ScreenWidth, sh := A_ScreenHeight
    gBar.Show(Format("x{} y{}", (sw-1080)//2, sh//6))
    gEdit.Focus()
}


; --- Context-sensitive keys: only when our GUI is active ---
#HotIf gBar && WinActive("ahk_id " gBar.Hwnd)
Enter::
{
    Submit()
    return
}
Esc::
{
    gBar.Hide()
    ToolTip()
    return
}
#HotIf

; --- Submit: build JSON and POST (async + WaitForResponse) ---
Submit() {
    global gBar, gEdit, API_URL, API_KEY
    title := Trim(gEdit.Value)
    if (title = "") {
        FlashTip("Please type a title first.")
        return
    }
    gBar.Hide()
    payload := BuildJsonWithTitle(title)

    if (DEBUG) {
        MsgBox("Payload to send:`n`n" payload)
    }


    try {
        WHR := ComObject("WinHttp.WinHttpRequest.5.1")
        WHR.Open("POST", API_URL, true) ; async
        WHR.SetRequestHeader("Content-Type", "application/json")
        WHR.SetRequestHeader("x-api-key", API_KEY)
; Fail fast instead of hanging
WHR.SetTimeouts(30000, 30000, 30000, 30000) ; resolve, connect, send, receive (ms)

        WHR.Send(payload)
        WHR.WaitForResponse() ; keep UI responsive
        status := WHR.Status
        resp   := WHR.ResponseText

        if (DEBUG) {
            MsgBox("HTTP Status: " status "`n`nResponse:`n`n" resp)
        }


        if (status >= 200 && status < 300) {
            FlashTip("Sent ✓ (" status ")", 1800)
        } else {
            FlashTip("Error " status ": " Shorten(resp), 4000)
        }
    } catch as e {
        if (DEBUG) {
            MsgBox("Request failed:`n" e.Message)
        }
        FlashTip("Request failed: " e.Message, 4000)
    }
}

; --- Helpers ---
BuildJsonWithTitle(title) {
    q := Chr(34) ; double quote
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
    ; Minimal JSON escape for quotes, backslash, and control chars
    q := Chr(34) ; "
    b := Chr(92) ; \
    out := ""
    for c in StrSplit(str, "") {
        code := Ord(c)
        if (c = q)               ; quote
            out .= b . q          ; -> \"
        else if (c = b)          ; backslash
            out .= b . b          ; -> \\
        else if (code = 8)       ; backspace
            out .= b "b"
        else if (code = 9)       ; tab
            out .= b "t"
        else if (code = 10)      ; newline
            out .= b "n"
        else if (code = 12)      ; form feed
            out .= b "f"
        else if (code = 13)      ; carriage return
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
    ; EM_SETCUEBANNER (0x1501): show placeholder text in an Edit control
    DllCall("user32\SendMessageW", "ptr", hEdit, "uint", 0x1501, "ptr", 1, "wstr", text)
}

SetEditMargins(hEdit, leftPx := 8, rightPx := 8) {
    ; EM_SETMARGINS (0x00D3): set left/right inner padding (MAKELONG)
    lParam := leftPx | (rightPx << 16)
    DllCall("user32\SendMessageW", "ptr", hEdit, "uint", 0x00D3, "ptr", 0x3, "ptr", lParam)
}

EnableDropShadow(hwnd) {
    ; Turn on non-client rendering so DWM draws a shadow for borderless windows
    DWMNCRP_ENABLED := 2
    try DllCall("dwmapi\DwmSetWindowAttribute","ptr",hwnd,"int",2,"int*",DWMNCRP_ENABLED,"int",4) ; DWMWA_NCRENDERING_POLICY=2
}

EnableAcrylic(hwnd, opacity := 175, tintRGB := 0x202020) {
    static WCA_ACCENT_POLICY := 19, ACCENT_ENABLE_ACRYLICBLURBEHIND := 4

    tintARGB := (opacity << 24) | (tintRGB & 0xFFFFFF)

    ; ACCENT_POLICY struct
    ACCENT := Buffer(16, 0)
    NumPut("Int", ACCENT_ENABLE_ACRYLICBLURBEHIND, ACCENT, 0)
    NumPut("Int", 0, ACCENT, 4)       ; Flags
    NumPut("Int", tintARGB, ACCENT, 8) ; GradientColor (ARGB)
    NumPut("Int", 0, ACCENT, 12)       ; AnimationId

    ; WINDOWCOMPOSITIONATTRIBUTEDATA struct
    DATA := Buffer(24, 0)
    NumPut("Int", WCA_ACCENT_POLICY, DATA, 0)   ; Attribute
    NumPut("Ptr", ACCENT.Ptr, DATA, 8)          ; Data
    NumPut("UInt", 16, DATA, 16)                ; Size of ACCENT

    try DllCall("user32\SetWindowCompositionAttribute", "ptr", hwnd, "ptr", DATA)
}

; ---- Win11Effects.ahk (AutoHotkey v2) ----
ApplyWin11Effects(hwnd, opts := Map(
    "Dark", true,              ; dark title bar
    "Corners", 2,              ; 2=Round
    "Backdrop", 2,             ; 2=Mica, 3=Transient (Acrylic-like), 4=Tabbed (Mica Alt)
    "Shadow", true             ; enable subtle DWM shadow
)) {
    ; Dark title bar
    if opts["Dark"] {
        val := 1
        DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "uint", 20, "int*", val, "uint", 4)
    }
    ; Rounded corners
    if opts.Has("Corners") {
        corner := Integer(opts["Corners"])
        DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "uint", 33, "int*", corner, "uint", 4)
    }
    ; System-drawn backdrop (Win11 22H2+)
    if opts.Has("Backdrop") {
        bt := Integer(opts["Backdrop"])  ; 2=Mica, 3=Transient (Acrylic-like), 4=Tabbed (Mica Alt)
        DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "uint", 38, "int*", bt, "uint", 4)
    }
    ; Subtle shadow on borderless windows: extend glass a hair into client area
    if opts["Shadow"] {
        ; MARGINS struct: left, right, top, bottom (4x Int)
        margins := Buffer(16, 0)
        NumPut("int", 1, margins, 0), NumPut("int", 1, margins, 4)
        NumPut("int", 1, margins, 8), NumPut("int", 1, margins, 12)
        DllCall("dwmapi\DwmExtendFrameIntoClientArea", "ptr", hwnd, "ptr", margins)
    }
}

