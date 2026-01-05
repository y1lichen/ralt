#Requires AutoHotkey v2.0
#SingleInstance Force
ListLines 0

; --- 全局變數與初始化 ---
Global SettingsFile := A_ScriptDir . "\ralt_settings.json"
Global StartupShortcut := A_Startup "\ralt.lnk"

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
    
    A_TrayMenu.Add("開機啟動", ToggleStartup)
    if FileExist(StartupShortcut)
        A_TrayMenu.Check("開機啟動")

    A_TrayMenu.Add()
    A_TrayMenu.Add("重新啟動", (*) => Reload())
    A_TrayMenu.Add("結束", (*) => ExitApp())
    A_TrayMenu.Default := "設定 ralt"
}

ToggleStartup(*) {
    if FileExist(StartupShortcut) {
        FileDelete(StartupShortcut)
        A_TrayMenu.Uncheck("開機啟動")
        MsgBox("已取消開機啟動", "ralt", "Iconi T3")
    } else {
        try {
            FileCreateShortcut(A_ScriptFullPath, StartupShortcut, A_ScriptDir)
            A_TrayMenu.Check("開機啟動")
            MsgBox("已設定開機時自動啟動", "ralt", "Iconi T3")
        } catch Error as err {
            MsgBox("設定失敗: " . err.Message, "錯誤", "Iconx")
        }
    }
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

    for id in ids {
        try {
            title := WinGetTitle(id)
            style := WinGetStyle(id)
            exStyle := WinGetExStyle(id)

            ; 1. 基礎過濾：必須有標題、可見、非工具視窗
            if (title == "" || !(style & 0x10000000) || (exStyle & 0x00000080))
                continue
                
            proc := WinGetProcessName(id)
            class := WinGetClass(id)
            
            ; 2. 判斷按鍵 (自定義優先，否則用首字母)
            appLetter := CustomMap.Has(proc) ? CustomMap[proc] : StrLower(SubStr(proc, 1, 1))
            
            if (appLetter == letter) {
                ; 3. 檔案總管過濾：只保留資料夾視窗
                if (proc == "explorer.exe" && class != "CabinetWClass")
                    continue
                
                ; 4. Edge 視窗過濾 (排除某些背景小組件)
                if (proc == "msedge.exe" && !InStr(title, "Edge"))
                    continue

                ; --- 核心修正：移除 uniqueApps 判斷 ---
                ; 現在同一個 App 的每個視窗都會被加入清單中
                AppGroups[letter].Push(id)
            }
        }
    }
}

RAlt::return