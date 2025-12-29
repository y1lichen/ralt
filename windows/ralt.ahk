#Requires AutoHotkey v2.0
#SingleInstance Force
ListLines 0

; --- 全局變數與初始化 ---
Global SettingsFile := A_ScriptDir . "\ralt_settings.json"
; 關鍵修正：設定 Map 為不分大小寫 (CaseSense Off)
Global CustomMap := Map()
CustomMap.CaseSense := "Off" 

Global AppGroups := Map()
Global LastKey := ""
Global CycleIndex := 0

LoadSettings()
SetupTray()
RegisterHotkeys()

; --- 1. Tray 選單設定 ---
SetupTray() {
    A_TrayMenu.Delete()
    A_TrayMenu.Add("設定 ralt", (*) => ShowSettingsGui())
    A_TrayMenu.Add()
    A_TrayMenu.Add("重新啟動", (*) => Reload())
    A_TrayMenu.Add("結束", (*) => ExitApp())
    A_TrayMenu.Default := "設定 ralt"
}

; --- 2. 設定視窗 GUI ---
ShowSettingsGui() {
    MyGui := Gui("+AlwaysOnTop", "ralt 自定義按鍵設定")
    MyGui.SetFont("s10", "Microsoft JhengHei")
    MyGui.Add("Text",, "目前開啟的 App (雙擊設定按鍵):")
    LV := MyGui.Add("ListView", "r15 w400", ["執行檔名", "目前對應鍵"])
    
    RefreshList(LV)
    
    BtnSave := MyGui.Add("Button", "w80 x160", "儲存並套用")
    BtnSave.OnEvent("Click", (*) => SaveAndApply(MyGui, LV))
    LV.OnEvent("DoubleClick", (LV, RowNumber) => SetKey(LV, RowNumber))
    MyGui.Show()
}

RefreshList(LV) {
    LV.Delete()
    ids := WinGetList(,, "Program Manager")
    seen := Map()
    seen.CaseSense := "Off"

    for id in ids {
        try {
            if !(WinGetStyle(id) & 0x10000000) || (WinGetExStyle(id) & 0x00000080)
                continue
            
            proc := WinGetProcessName(id)
            if seen.Has(proc)
                continue
            seen[proc] := true
            
            currentKey := CustomMap.Has(proc) ? CustomMap[proc] : ""
            LV.Add(, proc, currentKey)
        }
    }
}

SetKey(LV, Row) {
    procName := LV.GetText(Row, 1)
    IB := InputBox("請輸入想要對應 " . procName . " 的按鍵:", "設定按鍵", "w250 h130")
    if IB.Result = "OK" {
        key := StrLower(SubStr(IB.Value, 1, 1))
        LV.Modify(Row,, procName, key)
    }
}

; --- 3. JSON 儲存與讀取 ---
SaveAndApply(GuiObj, LV) {
    Global CustomMap := Map()
    CustomMap.CaseSense := "Off"
    
    jsonStr := "{"
    Loop LV.GetCount() {
        proc := LV.GetText(A_Index, 1)
        key := LV.GetText(A_Index, 2)
        if (key != "") {
            CustomMap[proc] := key
            jsonStr .= "`n  `"" . proc . "`": `"" . key . "`","
        }
    }
    jsonStr := RTrim(jsonStr, ",") . "`n}"
    
    if FileExist(SettingsFile)
        FileDelete(SettingsFile)
    FileAppend(jsonStr, SettingsFile, "UTF-8")
    
    GuiObj.Destroy()
    Reload()
}

LoadSettings() {
    if !FileExist(SettingsFile)
        return
    
    content := FileRead(SettingsFile)
    pos := 1
    while RegExMatch(content, "`"(.+?)`":\s*`"(.+?)`"", &match, pos) {
        CustomMap[match[1]] := match[2]
        pos := match.Pos + match.Len
    }
}

; --- 4. 核心切換邏輯 ---
RegisterHotkeys() {
    Loop 26 {
        char := Chr(96 + A_Index)
        Hotkey("RAlt & " . char, HandleSwitch)
    }
}

HandleSwitch(HotkeyName) {
    Global LastKey, CycleIndex, AppGroups
    currentKey := SubStr(HotkeyName, -1)
    
    if (currentKey != LastKey) {
        CycleIndex := 1
        LastKey := currentKey
        RefreshAppGroup(currentKey)
    } else {
        CycleIndex++
    }
    
    if !AppGroups.Has(currentKey) || AppGroups[currentKey].Length == 0
        return
    
    targetList := AppGroups[currentKey]
    if (CycleIndex > targetList.Length)
        CycleIndex := 1
    
    targetID := targetList[CycleIndex]
    if WinExist("ahk_id " . targetID) {
        WinActivate("ahk_id " . targetID)
    }
}

RefreshAppGroup(letter) {
    Global AppGroups, CustomMap
    AppGroups[letter] := []
    ids := WinGetList(,, "Program Manager")
    uniqueApps := Map()
    uniqueApps.CaseSense := "Off"

    for id in ids {
        try {
            title := WinGetTitle(id)
            style := WinGetStyle(id)
            exStyle := WinGetExStyle(id)

            ; 1. 基本過濾：必須有標題、必須可見、不能是工具視窗
            if (title == "" || !(style & 0x10000000) || (exStyle & 0x00000080))
                continue
                
            proc := WinGetProcessName(id)
            class := WinGetClass(id)
            
            ; 2. 決定按鍵 (CustomMap 優先)
            appLetter := CustomMap.Has(proc) ? CustomMap[proc] : StrLower(SubStr(proc, 1, 1))
            
            if (appLetter == letter) {
                ; 3. 特殊處理檔案總管
                if (proc == "explorer.exe" && class != "CabinetWClass")
                    continue
                
                ; 4. 特殊處理 Edge (排除那些沒有視窗主體的背景進程)
                if (proc == "msedge.exe" && !InStr(title, "Microsoft​ Edge")) {
                    ; 有些 Edge 視窗是 PWA 或背景，我們只取標題包含 Edge 的
                    ; 如果是 PWA 視窗，這行可以根據需求調整
                }

                if !uniqueApps.Has(proc) {
                    uniqueApps[proc] := id
                    AppGroups[letter].Push(id)
                }
            }
        }
    }
}

RAlt::return