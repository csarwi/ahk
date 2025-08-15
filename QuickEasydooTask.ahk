#Requires AutoHotkey v2.0

; ====== CONFIG ======
API_URL := "https://api.easydoo.com/api/integration/CreateOrUpdateEntity"
API_KEY := EnvGet("EASYDOO_API_KEY")
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
gBar := Gui("+AlwaysOnTop -Caption +ToolWindow +Border", "Quick POST")

	      ; bigger + modern spacing
	      gBar.MarginX := 24
	      gBar.MarginY := 18
	      gBar.SetFont("s20", "Segoe UI")  ; or "Segoe UI Variable" on Win11

	      ; wider, taller input
	      gEdit := gBar.Add("Edit", "w1000 h52 -Wrap -VScroll -HScroll +0x200")  ; +0x200 = ES_AUTOHSCROLL
	      gEdit.SetFont("s20")
	      gEdit.OnEvent("Focus", (*) => Send("^a"))

	      ; placeholder + inner padding
	      SetCueBanner(gEdit.Hwnd, "Please enter Name for easydoo-Workitem ...")
	      SetEditMargins(gEdit.Hwnd, 14, 14)

	      ; slight transparency (nice on dark/light)
	      try WinSetTransparent(235, "ahk_id " gBar.Hwnd)

	      ; subtle rounded corners (Win11+)
	      try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", gBar.Hwnd, "int", 33, "int*", 2, "int", 4)

	    }

    gEdit.Value := ""
	sw := A_ScreenWidth, sh := A_ScreenHeight
gBar.Show(Format("x{} y{}", (sw-1048)//2, sh//6))  ; 1000 + margins ≈ 1048
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
    gBar.Hide()
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

