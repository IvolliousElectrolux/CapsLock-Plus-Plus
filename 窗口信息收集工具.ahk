#Requires AutoHotkey v2.0

; 窗口信息收集工具 (增强版)
; 使用方法: 按下 Ctrl+Win+Alt+I 获取鼠标所在窗口及其控件信息

; 定义快捷键：Ctrl+Win+Alt+I
^#!i::GetWindowInfo()

GetWindowInfo() {
    ; 获取鼠标位置的窗口及控件
    MouseGetPos(&mouseX, &mouseY, &mouseWin, &mouseControl)
    if !mouseWin {
        MsgBox("无法获取鼠标下的窗口！")
        return
    }
    
    ; 收集窗口信息
    hwnd := mouseWin
    title := WinGetTitle("ahk_id " hwnd)
    class := WinGetClass("ahk_id " hwnd)
    pid := WinGetPID("ahk_id " hwnd)
    processName := WinGetProcessName("ahk_id " hwnd)
    style := WinGetStyle("ahk_id " hwnd)
    exStyle := WinGetExStyle("ahk_id " hwnd)
    
    try {
        processPath := ProcessGetPath(pid)
    } catch {
        processPath := "无法获取进程路径"
    }
    
    ; 获取窗口位置和大小
    WinGetPos(&winX, &winY, &winWidth, &winHeight, "ahk_id " hwnd)
    
    ; 检查窗口状态
    isVisible := style & 0x10000000 ? "是" : "否"  ; WS_VISIBLE
    isToolWindow := exStyle & 0x80 ? "是" : "否"  ; WS_EX_TOOLWINDOW
    isChildWindow := style & 0x40000000 ? "是" : "否"  ; WS_CHILD
    
    ; 准备信息显示
    info := "窗口信息:`n`n"
    info .= "窗口句柄 (HWND): " hwnd "`n"
    info .= "窗口标题: " title "`n"
    info .= "窗口类名: " class "`n"
    info .= "窗口位置: X=" winX ", Y=" winY ", 宽=" winWidth ", 高=" winHeight "`n`n"
    
    info .= "进程 ID (PID): " pid "`n"
    info .= "进程名称: " processName "`n"
    info .= "进程路径: " processPath "`n`n"
    
    info .= "窗口样式: 0x" Format("{:X}", style) "`n"
    info .= "窗口扩展样式: 0x" Format("{:X}", exStyle) "`n`n"
    
    info .= "可见窗口: " isVisible "`n"
    info .= "工具窗口: " isToolWindow "`n"
    info .= "子窗口: " isChildWindow "`n`n"
    
    ; 收集控件信息
    info .= "=== 控件信息 ===`n`n"
    if (mouseControl) {
        info .= "鼠标下的控件: " mouseControl "`n"
        
        ; 获取控件位置和大小
        try {
            ControlGetPos(&ctrlX, &ctrlY, &ctrlWidth, &ctrlHeight, mouseControl, "ahk_id " hwnd)
            info .= "控件位置: X=" ctrlX ", Y=" ctrlY ", 宽=" ctrlWidth ", 高=" ctrlHeight "`n"
        } catch {
            info .= "无法获取控件位置和大小`n"
        }
        
        ; 获取控件文本
        try {
            ctrlText := ControlGetText(mouseControl, "ahk_id " hwnd)
            info .= "控件文本: " ctrlText "`n"
        } catch {
            info .= "无法获取控件文本`n"
        }
        
        ; 获取控件句柄
        try {
            ctrlHwnd := ControlGetHwnd(mouseControl, "ahk_id " hwnd)
            info .= "控件句柄: " ctrlHwnd "`n"
            
            ; 获取控件样式
            ctrlStyle := SendMessage(0x0130, 0, 0, mouseControl, "ahk_id " hwnd)  ; WM_GETCONTROLSTYLE
            info .= "控件样式: 0x" Format("{:X}", ctrlStyle) "`n"
        } catch {
            info .= "无法获取控件句柄`n"
        }
    } else {
        info .= "鼠标下没有检测到控件`n"
    }
    
    ; 列出窗口中的所有控件
    info .= "`n=== 窗口内所有控件 ===`n`n"
    try {
        controls := WinGetControls("ahk_id " hwnd)
        if (controls.Length > 0) {
            for index, control in controls {
                info .= index ": " control "`n"
                
                ; 选择性地为重要控件显示更多信息
                if (index <= 10) {  ; 只对前10个控件获取详细信息以避免信息过多
                    try {
                        ctrlText := ControlGetText(control, "ahk_id " hwnd)
                        if (ctrlText)
                            info .= "   文本: " ctrlText "`n"
                        
                        ControlGetPos(&ctrlX, &ctrlY, &ctrlWidth, &ctrlHeight, control, "ahk_id " hwnd)
                        info .= "   位置: X=" ctrlX ", Y=" ctrlY ", 宽=" ctrlWidth ", 高=" ctrlHeight "`n"
                    } catch {
                        ; 忽略错误
                    }
                }
            }
        } else {
            info .= "未检测到控件`n"
        }
    } catch {
        info .= "无法获取控件列表`n"
    }
    
    info .= "`n=== 代码参考 ===`n"
    info .= "- 窗口标识: `n  `"ahk_class " class "`"`n"
    info .= "- 按类名排除: `n  `"" class "`",`n"
    info .= "- 按进程路径排除: `n  `"" processPath "`",`n"
    
    if (mouseControl) {
        info .= "- 控件操作示例: `n"
        info .= "  ControlClick(`"" mouseControl "`", `"ahk_id " hwnd "`")`n"
        info .= "  ControlSetText(`"新文本`", `"" mouseControl "`", `"ahk_id " hwnd "`")`n"
        info .= "  ControlGetText(`"" mouseControl "`", `"ahk_id " hwnd "`")`n"
    }
    
    ; 显示收集到的信息
    MsgBox(info, "窗口和控件信息", "Owner T4096")
    
    ; 复制到剪贴板
    A_Clipboard := info
}